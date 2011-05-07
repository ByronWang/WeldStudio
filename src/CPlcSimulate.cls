VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlcSimulate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IServer

Dim sDef As String
Dim Paragram As String
Dim from As Integer
Dim length As Integer

Private ra() As PlcDeclare.Record
Private index As Integer
Dim opened As Boolean

Dim defines(1000) As String
Dim define_index As Integer

'Function Prototypes
Public Function IServer_Init(ByVal def&) As Long
    IServer_Init = PlcDeclare.DTL_SUCCESS
End Function
Public Sub IServer_Uninit(ByVal param As Long)
    sDef = ""
End Sub
 '"N22:0,12,WORD,READ,AB:LOCAL,0,SLC500,1")

Public Function IServer_Define(ByRef NameId&, ByVal def As String) As Long
    Dim i As Integer
    For i = 0 To define_index
        If defines(i) = def Then
            NameId = i
            Exit Function
        End If
    Next

    define_index = define_index + 1
    defines(define_index) = def
    NameId = define_index
    IServer_Define = PlcDeclare.DTL_SUCCESS
End Function

Public Function IServer_Undef(ByVal NameId&) As Long
    IServer_Undef = PlcDeclare.DTL_SUCCESS
End Function

Public Function IServer_OpenDriver(ByVal nDriverId&, ByVal szDriverName As String, ByVal timeout&) As Long

    
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject
    
    If fso.FileExists(PLCDrv.SimulatePath) And FileLen(PLCDrv.SimulatePath) > 512 Then
        IServer_OpenDriver = PlcDeclare.DTL_SUCCESS
    Else
        IServer_OpenDriver = PlcDeclare.DTL_E_FAIL
    End If
    
End Function
Public Function IServer_CloseDriver(ByVal nDriverId&, ByVal timeout&) As Long
    IServer_CloseDriver = PlcDeclare.DTL_SUCCESS
End Function

Public Function IServer_ReadInt(ByVal NameId&, ByRef Variable() As Integer, Iostat&, ByVal timeout&) As Long
    Dim def As String
    
    def = defines(NameId&)
    
    Dim sa() As String
    sa = Split(def, ",")

    Dim saa() As String
    saa = Split(sa(0), ":")
    from = CInt(saa(1))
    length = CInt(sa(1))
    Paragram = saa(0)
    
    
    Select Case Paragram
        Case "N22":
            If opened Then
                If 0 <= index And index < UBound(ra) Then
                    Dim r As PlcDeclare.Record
                    
                    r = ra(index)
                       
                    Variable(0) = r.data.Dist * 100
                    Variable(1) = r.data.Amp
                    Variable(2) = r.data.PsiUpset
                    Variable(3) = r.data.Volt
                    Variable(4) = r.data.PsiOpen
                    Variable(8) = 1
                    Variable(9) = r.data.WeldStage
                    Variable(11) = r.data.PlcStage
                    index = index + 1
                ElseIf index < UBound(ra) + 100 Then
                    Variable(0) = 10 * 100
                    Variable(1) = 17
                    Variable(2) = 98
                    Variable(3) = 90
                    Variable(4) = 90
                    Variable(9) = -1
                    Variable(11) = -1
                    index = index + 1
                Else
                    If PLCDrv.SimulatePath = "" Then
                        Variable(0) = 10 * 100
                        Variable(1) = 17
                        Variable(2) = 98
                        Variable(3) = 90
                        Variable(4) = 90
                        Variable(8) = 0
                        Variable(9) = 90
                        Variable(11) = 90
                    Else
                        Variable(0) = 10 * 100
                        Variable(1) = 17
                        Variable(2) = 98
                        Variable(3) = 90
                        Variable(4) = 90
                        Variable(8) = 0
                        Variable(9) = -1
                        Variable(11) = -1
                        ra = PlcWld.LoadData(PLCDrv.SimulatePath).data
                        opened = True
                        index = 0
                    End If
                End If
            Else
                    If PLCDrv.SimulatePath = "" Then
                        Variable(0) = 10 * 100
                        Variable(1) = 17
                        Variable(2) = 98
                        Variable(3) = 90
                        Variable(4) = 90
                        Variable(8) = 0
                        Variable(9) = 90
                        Variable(11) = 90
                    Else
                        ra = PlcWld.LoadData(PLCDrv.SimulatePath).data
                        index = 0
                        opened = True
                    End If
            End If
        Case "N21":
            If from = "3" Then
                Variable(0) = PULSE_MODE
            End If
    End Select
   ' IServer_ReadInt = PlcDeclare.DTL_READ_W(NameId, Variable(0), Iostat, timeout)
    IServer_ReadInt = PlcDeclare.DTL_SUCCESS
End Function

Public Function IServer_ReadSingle(ByVal NameId&, ByRef Variable() As Single, Iostat&, ByVal timeout&) As Long
        Dim def As String
    
    def = defines(NameId&)
    
    Dim sa() As String
    sa = Split(def, ",")

    Dim saa() As String
    saa = Split(sa(0), ":")
    from = CInt(saa(1))
    length = CInt(sa(1))
    Paragram = saa(0)
    
    
    Select Case Paragram
        Case "F62":
            Variable(0) = 10 * 100
            Variable(1) = 17
            Variable(2) = 98
            Variable(3) = 90
            Variable(4) = 90
            Variable(8) = 0
            Variable(9) = 90
            Variable(11) = 90
                
    End Select
    
    IServer_ReadSingle = PlcDeclare.DTL_SUCCESS
End Function


Public Function IServer_WriteInt(ByVal NameId&, ByRef Variable() As Integer, Iostat&, ByVal timeout&) As Long
    IServer_WriteInt = PlcDeclare.DTL_SUCCESS
End Function

Public Function IServer_WriteSingle(ByVal NameId&, Variable() As Single, Iostat&, ByVal timeout&) As Long
    IServer_WriteSingle = PlcDeclare.DTL_SUCCESS
End Function

Public Sub IServer_ErrorStr(ByVal status&, ByVal errstr$, ByVal StrSize%)
End Sub



