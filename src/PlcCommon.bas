Attribute VB_Name = "out"

Public Sub log(msg As String)

    Dim fn As String
    fn = App.path & "\" & Format(Date, "YYYY-MM-DD") & ".log"
    Open fn For Append As #5
    Print #5, Date & " " & Time & " -- " & msg
    Close #5

End Sub



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



