VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EngWeldNumber"
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

Private Function IWeldNumber_Compacted() As String
    IWeldNumber_Compacted = base.Compacted
End Function

Private Function IWeldNumber_Display() As String
    Dim unit As String
    Dim operator As String
    
    unit = GetSetting(App.EXEName, "UserData", "Unit", "")
    operator = GetSetting(App.EXEName, "UserData", "Operator", "")

    IWeldNumber_Display = unit & base.Display & operator
End Function

Private Function IWeldNumber_MoveNext() As Variant
    base.MoveNext
End Function

Private Function IWeldNumber_Update(strWeldNumber As String) As Boolean
    IWeldNumber_Update = base.Update(strWeldNumber)
End Function

Private Function IWeldNumber_FromCompact(s As String) As Long
    IWeldNumber_FromCompact = base.FromCompact(s)
End Function

Private Function IWeldNumber_ToCompact(n As Long) As String
    IWeldNumber_ToCompact = base.ToCompact(n)
End Function

Private Function IWeldNumber_FromDisplay(sFull As String) As Long
    IWeldNumber_FromDisplay = base.FromDisplay(sFull)
End Function

Private Function IWeldNumber_ToDisplay(dat As Date, n As Long) As String
    IWeldNumber_ToDisplay = base.ToDisplay(dat, n)
End Function

