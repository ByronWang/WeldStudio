Attribute VB_Name = "PlcRes"
Option Explicit

Public LANGUAGE As Integer



Private Sub Init()
Dim l As String
LANGUAGE = -1

    l = GetSetting(App.EXEName, "General", "Language", "ZH")

    If l = "ZH" Then
        LANGUAGE = 0
    End If
End Sub

Public Sub LoadResFor(frm As Form)
Init

Dim c As Control
Dim bas As Long
bas = LANGUAGE

If LANGUAGE = 0 Then

    
If frm.Tag <> "" Then
    frm.Caption = LoadResString(bas + CInt(frm.Tag)) '& "="
End If

    Dim i As Integer

Dim v As Integer
    For Each c In frm.Controls
        If c.Tag = "" Then
            'SetIndex c, 60000
        ElseIf Left(c.name, 3) = "mnu" Then
            SetString c, bas + CInt(c.Tag)
        ElseIf Left(c.name, 5) <> "Frame" And Left(c.Container.name, 5) = "Frame" Then
           v = bas + CInt(c.Container.Tag) + CInt(c.Tag)
            Call SetString(c, v)
        Else
            SetString c, bas + CInt(c.Tag)
        End If
         
    Next
End If
End Sub

Public Function LoadMsgResString(id As Long) As String
On Error Resume Next
    LoadMsgResString = LoadResString(LANGUAGE + id)
End Function

Private Sub SetContranerRes(index As Integer, bas As Integer, con As Control)
Dim c As Control
On Error Resume Next

    Dim i As Integer
    For Each c In con.Container.Controls
        i = i + 1
        If c.Tag <> "" Then
            SetString c, bas + CInt(c.Tag)
        Else
            SetIndex c, index * 100 + i
        End If
        'Call SetContranerRes(bas + CInt(c.Tag), c)
    Next
End Sub


Private Sub SetString(c As Control, id As Integer)

Dim s As String
    If id >= 0 Then
    On Error Resume Next
        s = LoadResString(id)
    On Error GoTo 0
    
        Debug.Print id & vbTab & s & vbTab & c.Caption
        
        If s <> "" Then
            c.Caption = s '& "="
        Else
           ' c.Caption
            'c.Caption = id & "=" & c.Caption
        End If
    Else
        c.Caption = (id + 30000) & "*" & c.Caption
    End If
End Sub

Private Sub SetIndex(c As Control, index As Integer)
    c.Caption = index & "." & c.Caption
End Sub
