Attribute VB_Name = "PlcCommon"



Public Function toWeldNumberShowModel(n As Integer) As String
Dim leadNumber As Integer
Dim leadChar As String
    
    leadNumber = (n / 10000)
    
    If leadNumber >= 26 Then
        leadNumber = 0
    End If
    
    leadChar = Chr(Asc("A") + leadNumber)
        
    Dim leaveNumber As Integer
    leaveNumber = n - CInt(n / 10000) * 10000
    
    Dim showString As String
    showString = CStr(leaveNumber)
    
    toWeldNumberShowModel = "" & leadChar & Left("0000", 4 - Len(showString)) & showString
End Function


Public Function fromWeldNumberShowModel(s As String) As Integer


Dim leadNumber As Integer
Dim leadChar As String
Dim leaveNumber As Integer
    
    leadChar = Left(s, 1)
    leadNumber = Asc(leadChar) - Asc("A")


    leaveNumber = CInt(Mid(s, 2, 4))
    
    fromWeldNumberShowModel = leadNumber * 10000 + leaveNumber
    
End Function



