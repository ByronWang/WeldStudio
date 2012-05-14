Attribute VB_Name = "PlcDailyReport"
Option Explicit

Public Type DailyReport
     Serial As String * 1
     X1 As Byte
     X2 As Byte
     X3 As Byte
     Sequence As Long
End Type


Public Function LoadData(FileName As String) As DailyReport()
Dim pos As Long
Dim i As Long
Dim r(500) As DailyReport
Dim dr As DailyReport

Open FileName For Binary As #1
    pos = 0
    i = 0
    
    While pos < LOF(1)
        Get 1, pos + 1, r(i)
        i = i + 1
        pos = pos + Len(dr)
    Wend
Close 1

Dim max As Long
max = i
Dim data() As DailyReport
ReDim data(max - 1)

For i = 1 To max
    data(i - 1) = r(i - 1)
Next i

LoadData = data

End Function



Public Function SaveData(FileName As String, data As DailyReport)
Dim pos As Long
Dim i As Long
Dim r(100) As DailyReport

Dim oldDaylyReport() As DailyReport

Dim fileExist As Boolean
fileExist = False
On Error GoTo HASNOT_FILE
oldDaylyReport = LoadData(FileName)
fileExist = True

HASNOT_FILE:
On Error GoTo 0


Open FileName For Binary As #1
    pos = 0
    i = 0
    If fileExist Then
    For i = LBound(oldDaylyReport) To UBound(oldDaylyReport)
        If oldDaylyReport(i).Serial <> data.Serial Or oldDaylyReport(i).Sequence <> data.Sequence Then
            Put 1, pos + 1, oldDaylyReport(i)
            pos = pos + Len(data)
        End If
    Next i
    End If
        
    Put 1, pos + 1, data
    pos = pos + Len(data)
    
Close 1

End Function


