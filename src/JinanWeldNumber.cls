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

Dim weldSerailNumber As Long

Dim base As IWeldNumber

Private Sub Class_Initialize()
    Set base = New GeneralWeldNumber
End Sub

Private Function IWeldNumber_Display() As String
    Dim unit As String
    Dim operator As String
    
    unit = GetSetting(App.EXEName, "UserData", "Unit", "")
    operator = GetSetting(App.EXEName, "UserData", "Operator", "")

    IWeldNumber_Display = unit & base.Display & operator
End Function

Private Function IWeldNumber_FileName() As String
    weldSerailNumber = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
    IWeldNumber_FileName = toWeldNumberShowModel(weldSerailNumber)
    
    
    IWeldNumber_FileName = base.FileName
End Function

Private Function IWeldNumber_LeaderChar() As String
    IWeldNumber_LeaderChar = base.LeaderChar
End Function

Private Function IWeldNumber_MoveNext() As Variant
    base.MoveNext
End Function

Private Function IWeldNumber_Sequence() As Integer
    base.MoveNext
End Function

Private Function IWeldNumber_Update(strWeldNumber As String) As Boolean
    IWeldNumber_Update = base.Update(strWeldNumber)
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






