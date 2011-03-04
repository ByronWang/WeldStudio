Attribute VB_Name = "PlcWld"
Option Explicit



Public Function LoadData(filename As String) As FileR
Dim pos As Long
Dim r() As Record
Dim fh1 As FileHeader1
Dim fh2 As FileHeader2

Dim analysisDefine As WeldAnalysisDefineType
Dim analysisResult As WeldAnalysisResultType


Open filename For Binary As #1

    pos = 0
    
    Get 1, pos + 1, fh1
    pos = pos + Len(fh1)
    
    '  ÎªÁË¼æÈÝÐÔ
    Dim fn As String * 5
    Get 1, pos + 1, fn
    If UCase(fn & ".wld") <> Right(UCase(filename), 9) Then
        pos = pos - 5
    End If
    
    
    Get 1, pos + 1, fh2
    pos = pos + Len(fh2)
    
    
    
    ReDim r(fh2.RecordCount - 1)
    Get 1, pos + 1, r
    pos = pos + Len(r(0)) * fh2.RecordCount
    
    pos = pos - 4
    Get 1, pos + 1, analysisDefine
    pos = pos + Len(analysisDefine)
    
    pos = pos + 40 'TODO Seperate
    
    Get 1, pos + 1, analysisResult
    pos = pos + Len(analysisResult)
    
    
Close 1

LoadData.data = r
LoadData.header1 = fh1
LoadData.header2 = fh2
LoadData.analysisDefine = analysisDefine
LoadData.analysisResult = analysisResult
End Function



Public Function SaveData(filename As String, fh1 As FileHeader1, fh2 As FileHeader2, data() As WeldData, count As Integer, _
        analysisDefine As WeldAnalysisDefineType, analysisResult As WeldAnalysisResultType)
Dim pos As Long
Dim r() As Record
ReDim r(count - 1)



Dim i As Integer

For i = 0 To count - 1
    r(i).data = data(i)
Next

fh2.RecordCount = count

Open filename For Binary As #1

    pos = 0
    Put 1, pos + 1, fh1
    pos = pos + Len(fh1)
    
    Put 1, pos + 1, fh2
    pos = pos + Len(fh2)
    
    Put 1, pos + 1, r
    pos = pos + Len(r(0)) * fh2.RecordCount
    
    pos = pos - 4
    Put 1, pos + 1, analysisDefine
    pos = pos + Len(analysisDefine)
    
    pos = pos + 40 'TODO Seperate
    
    Put 1, pos + 1, analysisResult
    pos = pos + Len(analysisResult)
Close 1

End Function


