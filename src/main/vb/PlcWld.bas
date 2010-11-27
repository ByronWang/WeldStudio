Attribute VB_Name = "PlcWld"
Option Explicit



Public Function LoadData(filename As String) As FileR
Dim pos As Long
Dim r() As Record
Dim fh As FileHeader

Dim analysisDefine As WeldAnalysisDefineType
Dim analysisResult As WeldAnalysisResultType


Open filename For Binary As #1

    pos = 0
    
    Get 1, pos + 1, fh
    pos = pos + Len(fh)
    ReDim r(fh.RecordCount - 1)
    Get 1, pos + 1, r
    pos = pos + Len(r(0)) * fh.RecordCount
    
    pos = pos - 4
    Get 1, pos + 1, analysisDefine
    pos = pos + Len(analysisDefine)
    
    pos = pos + 40 'TODO Seperate
    
    Get 1, pos + 1, analysisResult
    pos = pos + Len(analysisResult)
    
    
Close 1

LoadData.data = r
LoadData.header = fh
LoadData.analysisDefine = analysisDefine
LoadData.analysisResult = analysisResult
End Function



Public Function SaveData(filename As String, fh As FileHeader, data() As WeldData, count As Integer, _
        analysisDefine As WeldAnalysisDefineType, analysisResult As WeldAnalysisResultType)
Dim pos As Long
Dim r() As Record
ReDim r(count - 1)



Dim i As Integer

For i = 0 To count - 1
    r(i).data = data(i)
Next

fh.RecordCount = count

Open filename For Binary As #1

    pos = 0
    Put 1, pos + 1, fh
    pos = pos + Len(fh)
    Put 1, pos + 1, r
    pos = pos + Len(r(0)) * fh.RecordCount
    
    pos = pos - 4
    Put 1, pos + 1, analysisDefine
    pos = pos + Len(analysisDefine)
    
    pos = pos + 40 'TODO Seperate
    
    Put 1, pos + 1, analysisResult
    pos = pos + Len(analysisResult)
Close 1

End Function


