Attribute VB_Name = "out"

Public Sub log(msg As String)
'
'    Dim fn As String
'    fn = App.path & "\" & Format(Date, "YYYY-MM-DD") & ".log"
'    Open fn For Append As #5
'    Print #5, Date & " " & Time & " -- " & msg
'    Close #5

End Sub


Public Sub logSingleArray(msg As String, a() As Single)
    out.log "   " & msg & " -> "
    Dim i As Integer
    For i = LBound(a) To UBound(a)
        out.log "         " & i & " : " & a(i)
    Next i
End Sub

Public Sub logIntegerArray(msg As String, a() As Integer)
    out.log "   " & msg & " -> "
    Dim i As Integer
    For i = LBound(a) To UBound(a)
        out.log "         " & i & " : " & a(i)
    Next i

End Sub

Public Function eq(x As Single, y As Single) As Boolean
    If x = y Then
        eq = True
    ElseIf Abs((x - y) / ((x + y) / 2)) < 1 / 50 And Abs(x - y) < 1 Then
        eq = True
    Else
        eq = False
    End If
End Function

