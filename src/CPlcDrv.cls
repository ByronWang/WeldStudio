VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlcDrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IServer

'Function Prototypes
Public Function IServer_Init(ByVal def&) As Long
    IServer_Init = PLCDrv.DTL_INIT(def)
End Function
Public Sub IServer_Uninit(ByVal param As Long)
    PLCDrv.DTL_UNINIT (param)
End Sub

Public Function IServer_Define(ByRef NameId&, ByVal def As String) As Long
    IServer_Define = PLCDrv.DTL_C_DEFINE(NameId, def)
End Function

Public Function IServer_Undef(ByVal NameId&) As Long
    IServer_Undef = PLCDrv.DTL_UNDEF(NameId&)
End Function

Public Function IServer_OpenDriver(ByVal nDriverId&, ByVal szDriverName As String, ByVal timeout&) As Long
    IServer_OpenDriver = PLCDrv.DTL_DRIVER_OPEN(nDriverId, szDriverName, timeout)
End Function
Public Function IServer_CloseDriver(ByVal nDriverId&, ByVal timeout&) As Long
    IServer_CloseDriver = PLCDrv.DTL_DRIVER_CLOSE(nDriverId, timeout)
End Function

Public Function IServer_ReadInt(ByVal NameId&, ByRef Variable() As Integer, Iostat&, ByVal timeout&) As Long
    IServer_ReadInt = PLCDrv.DTL_READ_W(NameId, Variable(0), Iostat, timeout)
End Function

Public Function IServer_ReadSingle(ByVal NameId&, ByRef Variable() As Single, Iostat&, ByVal timeout&) As Long
    IServer_ReadSingle = PLCDrv.DTL_READ_W(NameId, Variable(0), Iostat, timeout)
End Function


Public Function IServer_WriteInt(ByVal NameId&, ByRef Variable() As Integer, Iostat&, ByVal timeout&) As Long
    IServer_WriteInt = PLCDrv.DTL_WRITE_W(NameId, Variable(0), Iostat, timeout)
End Function

Public Function IServer_WriteSingle(ByVal NameId&, ByRef Variable() As Single, Iostat&, ByVal timeout&) As Long
    IServer_WriteSingle = PLCDrv.DTL_WRITE_W(NameId, Variable(0), Iostat, timeout)
End Function

Public Sub IServer_ErrorStr(ByVal Status&, ByVal errstr$, ByVal StrSize%)
     PLCDrv.DTL_ERROR_S Status, errstr$, StrSize
End Sub


