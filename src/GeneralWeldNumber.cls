VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeneralWeldNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWeldNumber

Dim weldSerailNumber As Long

Private Sub Class_Initialize()
    weldSerailNumber = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
End Sub

Private Function IWeldNumber_Compacted() As String
    IWeldNumber_Compacted = IWeldNumber_ToCompact(weldSerailNumber)
End Function

Private Function IWeldNumber_Display() As String
    IWeldNumber_Display = IWeldNumber_Compacted
End Function

Private Function IWeldNumber_MoveNext() As Variant
    Dim wn As Long
    wn = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
    If wn = weldSerailNumber Then
        weldSerailNumber = weldSerailNumber + 1
        Call SaveSetting(App.EXEName, "WELD", "LastSerialNumber", weldSerailNumber)
    End If
End Function

Private Function IWeldNumber_Update(str As String) As Boolean
    If Len(str) <> 5 Then
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
    
    leadChar = Left(s, 1)
    leadNumber = Asc(leadChar) - Asc("A")

    leaveNumber = CInt(Mid(s, 2, 4))
    
    IWeldNumber_FromCompact = leadNumber * 10000 + leaveNumber
End Function

Private Function IWeldNumber_ToCompact(n As Long) As String
    Dim leadNumber As Long
    Dim leadChar As String
    
    leadNumber = (n / 10000)
    
    If leadNumber >= 26 Then
        leadNumber = 0
    End If
    
    leadChar = Chr(Asc("A") + leadNumber)
    
    Dim leaveNumber As Long
    leaveNumber = n - 10000# * CInt(n / 10000)
    
    Dim showString As String
    showString = CStr(leaveNumber)
    
    IWeldNumber_ToCompact = "" & leadChar & Left("0000", 4 - Len(showString)) & showString
End Function

Private Function IWeldNumber_FromDisplay(sFull As String) As Long
    IWeldNumber_FromDisplay = IWeldNumber_FromCompact(sFull)
End Function

Private Function IWeldNumber_ToDisplay(dat As Date, n As Long) As String
    IWeldNumber_ToDisplay = IWeldNumber_ToCompact(n)
End Function
