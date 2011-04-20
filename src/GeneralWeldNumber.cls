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
Private Function IWeldNumber_Display() As String
    IWeldNumber_Display = IWeldNumber_FileName
End Function

Private Function IWeldNumber_FileName() As String
    weldSerailNumber = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
    IWeldNumber_FileName = toWeldNumberShowModel(weldSerailNumber)
End Function

Private Function IWeldNumber_LeaderChar() As String
    Dim sn As String
    sn = IWeldNumber_FileName
    IWeldNumber_LeaderChar = Left(sn, 1)
End Function

Private Function IWeldNumber_MoveNext() As Variant
    Dim wn As Integer
    wn = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
    If wn = weldSerailNumber Then
        weldSerailNumber = weldSerailNumber + 1
        Call SaveSetting(App.EXEName, "WELD", "LastSerialNumber", weldSerailNumber)
    End If
End Function

Private Function IWeldNumber_Sequence() As Integer
    Dim sn As String
    sn = IWeldNumber_FileName
    IWeldNumber_Sequence = CInt(Mid(sn, 2))
End Function

Private Function IWeldNumber_Update(strWeldNumber As String) As Boolean
    If Len(strWeldNumber) <> 5 Then
        IWeldNumber_Update = False
        Exit Function
    End If
    
   Dim weldSerailNumber As Long
   weldSerailNumber = fromWeldNumberShowModel(strWeldNumber)
   Call SaveSetting(App.EXEName, "WELD", "LastSerialNumber", weldSerailNumber)
   IWeldNumber_Update = True
End Function


Public Function toWeldNumberShowModel(n As Long) As String
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
    
    toWeldNumberShowModel = "" & leadChar & Left("0000", 4 - Len(showString)) & showString
End Function


Public Function fromWeldNumberShowModel(s As String) As Long


Dim leadNumber As Long
Dim leadChar As String
Dim leaveNumber As Long
    
    leadChar = UCase(Left(s, 1))
    leadNumber = Asc(leadChar) - Asc("A")


    leaveNumber = CInt(Mid(s, 2, 4))
    
    fromWeldNumberShowModel = leadNumber * 10000 + leaveNumber
    
End Function





