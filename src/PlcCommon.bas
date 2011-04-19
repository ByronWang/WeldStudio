Attribute VB_Name = "out"

Public Sub log(msg As String)

    Dim fn As String
    fn = App.path & "\" & Format(Date, "YYYY-MM-DD") & ".log"
    Open fn For Append As #5
    Print #5, Date & " " & Time & " -- " & msg
    Close #5

End Sub

