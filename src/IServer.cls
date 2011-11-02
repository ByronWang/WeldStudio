VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Init(ByVal def&) As Long

End Function
Public Sub Uninit(ByVal param As Long)

End Sub

Public Function Define(ByRef NameId&, ByVal def As String) As Long

End Function

Public Function Undef(ByVal NameId&) As Long

End Function

Public Function OpenDriver(ByVal nDriverId&, ByVal szDriverName As String, ByVal timeout&) As Long

End Function
Public Function CloseDriver(ByVal nDriverId&, ByVal timeout&) As Long

End Function

Public Function ReadInt(ByVal NameId&, ByRef Variable() As Integer, Iostat&, ByVal timeout&) As Long

End Function

Public Function ReadSingle(ByVal NameId&, ByRef Variable() As Single, Iostat&, ByVal timeout&) As Long

End Function


Public Function WriteInt(ByVal NameId&, ByRef Variable() As Integer, Iostat&, ByVal timeout&) As Long

End Function

Public Function WriteSingle(ByVal NameId&, ByRef Variable() As Single, Iostat&, ByVal timeout&) As Long

End Function

Public Sub ErrorStr(ByVal Status&, ByVal errstr$, ByVal StrSize%)

End Sub



