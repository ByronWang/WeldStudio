VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   7965
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Text            =   "C:\WeldStudio\WeldAid\luwencool\13-Jun-2009\A0007.WLD"
      Top             =   600
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00D2FB73&
      Caption         =   "Command1"
      Height          =   1455
      Left            =   1440
      TabIndex        =   0
      Top             =   1920
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Type rt
    s As String * 12
    d As Single
End Type


Private Sub Command1_Click()
read
End Sub

Private Function bw()

Dim filename As String

filename = "c:\twe.txt"

Dim pos As Integer
pos = 0

Dim s As Single
Dim r As rt



Dim b As Boolean

b = True


Open filename For Binary As #1

   wr pos, 0.04
   pos = pos + Len(r)
   
   wr pos, 0.45
   pos = pos + Len(r)
   
   wr pos, 0.45
   pos = pos + Len(r)
   
   
   wr pos, 3.2
   pos = pos + Len(r)
   
   wr pos, 8#
   pos = pos + Len(r)
      
   wr pos, 20#
   pos = pos + Len(r)
   
   


Close 1



End Function

Private Function writfe()

Dim filename As String

filename = "c:\twe.txt"

Dim pos As Integer
pos = 0

Dim s As Single
Dim r As rt


Open filename For Binary As #1

   wr pos, 0.04
   pos = pos + Len(r)
   
   wr pos, 0.45
   pos = pos + Len(r)
   
   wr pos, 0.45
   pos = pos + Len(r)
   
   
   wr pos, 3.2
   pos = pos + Len(r)
   
   wr pos, 8#
   pos = pos + Len(r)
      
   wr pos, 20#
   pos = pos + Len(r)
   
   


Close 1


End Function



Private Function read()

Dim filename As String

filename = Me.Text1.Text

Dim s As Single
Dim d As Long

Dim i As Integer

Dim c As Integer


Dim pos As Long
Dim r() As Record
Dim fh As FileHeader




Open filename For Binary As #1



Dim phPos As Long


    pos = 0
    
    Get 1, pos + 1, fh
    pos = pos + Len(fh)
    ReDim r(fh.RecordCount - 1)
    Get 1, pos + 1, r

    pos = pos + Len(r(0)) * fh.RecordCount
phPos = pos



Dim w As WeldAnalysisType
Dim inn As InnerAna
Dim a As AnaResult

pos = phPos
   Get 1, pos + 1 - 4, w
   pos = pos - 4 + Len(w)
   Get 1, pos + 1, inn
   pos = pos + Len(inn)
   Get 1, pos + 1, a
   pos = pos + Len(a)
   
   Debug.Print d
Dim b As Boolean


c = &H1AB / 4

pos = phPos
For i = 0 To c
   Get 1, pos + 1, s
   pos = pos + Len(s)
   Debug.Print s
Next i
    
pos = phPos
For i = 0 To c
   Get 1, pos + 1, d
   pos = pos + Len(d)
   Debug.Print d
Next

Close 1


End Function


Private Function wr(pos As Integer, d As Single)
    Dim r As rt
    
    r.s = d
    r.d = d
    
    Put 1, pos + 1, r
    
End Function




