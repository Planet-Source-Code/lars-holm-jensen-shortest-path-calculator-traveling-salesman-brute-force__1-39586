VERSION 5.00
Begin VB.Form frmTSP 
   Caption         =   "Shortest Path Calculator - (Brute Force!)"
   ClientHeight    =   7032
   ClientLeft      =   2460
   ClientTop       =   2088
   ClientWidth     =   8112
   LinkTopic       =   "Form1"
   ScaleHeight     =   586
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   676
   Begin VB.CheckBox chkAutoDraw 
      Caption         =   "AutoDraw"
      Height          =   252
      Left            =   6960
      TabIndex        =   7
      ToolTipText     =   "Draw shortest path while dragging"
      Top             =   120
      Value           =   1  'Checked
      Width           =   972
   End
   Begin VB.CheckBox chkDist 
      Caption         =   "Show Distances"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   120
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkEdges 
      Caption         =   "Show Vertices"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   1332
   End
   Begin VB.CommandButton cmdShortPath 
      Caption         =   "Shortest Path"
      Height          =   252
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdNed 
      Height          =   132
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   132
   End
   Begin VB.CommandButton cmdOp 
      Height          =   132
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   132
   End
   Begin VB.TextBox txtVertexes 
      Height          =   288
      Left            =   1440
      TabIndex        =   0
      Text            =   "5"
      Top             =   120
      Width           =   972
   End
   Begin VB.Label lblNumber 
      Caption         =   "Number of points:"
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   45
   End
End
Attribute VB_Name = "frmTSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

'*******************************'
'*   TSP-calc is written by    *'
'*  Lars Holm Jensen, Denmark  *'
'* larsholmjensen@hotmail.com  *'
'*******************************'

' The grab handle we are dragging. This is 0
' when we are not dragging any handle.
Private m_DraggingHandle As Integer, m_StartHandle As Integer, m_EndHandle As Integer

' The data points.
Private m_NumPoints As Single
Private m_PointX() As Single
Private m_PointY() As Single

'Shortest Path Calculation
    Private Distances() As Single, DrawShortest As Boolean
    Private InShortest()
    
    'Permutation
    Private NumValues As Integer, CurSolNum As Long
    Private Used() As Integer
    Private CurrentSolution() As Integer
'Shortest... end

Private Const HANDLE_WIDTH = 6
Private Const HANDLE_HALF_WIDTH = HANDLE_WIDTH / 2
Private Const TwoPi = 6.28318530717959

Private Type POINTAPI
        x As Long
        y As Long
End Type

Dim Cross() As POINTAPI
Dim NumCross As Integer, ReelNumCross As Integer

Dim EdgeColor As Long
Dim ShortEdgeColor As Long
Dim VertexColor As Long
Dim VertexBorderColor As Long
Dim StartVertexColor As Long
Dim EndVertexColor As Long
Dim TextColor As Long



Private Sub chkDist_Click()
If chkDist.Value = 1 Then TextColor = vbBlue Else TextColor = Me.BackColor
Refresh

End Sub

Private Sub chkEdges_Click()
If chkEdges.Value = 1 Then EdgeColor = vbBlack Else EdgeColor = Me.BackColor
Refresh

End Sub

Private Sub cmdNed_Click()
If Val(txtVertexes.Text) > 3 Then txtVertexes.Text = Val(txtVertexes.Text) - 1
Vertex_Change (Val(txtVertexes.Text))

End Sub

Private Sub cmdOp_Click()
txtVertexes.Text = Val(txtVertexes.Text) + 1
Vertex_Change (Val(txtVertexes.Text))

End Sub

Private Sub cmdShortPath_Click()
' If ending points not selected exit sub
'If (m_StartHandle <> 0 And m_EndHandle <> 0) Then CalcShortest
tid = Timer
CalcShortest
tid = Timer - tid
If tid > 2 Then tid = Int(tid)
Me.Caption = "Shortest Path Calculator - Points: " & m_NumPoints & ", Time used: " & tid & " sec"
End Sub

Private Sub CalcShortest()
ReDim Distances(1 To m_NumPoints, 1 To m_NumPoints)
Dim i As Integer, j As Integer
Dim MinimumLength As Single, Route As Long, MySum As Single, PointNum As Integer
Dim ShortestPath As Long

Label1.Caption = ""
DoEvents

If m_NumPoints < 3 Then Exit Sub

' Calculate distances..
For i = 1 To m_NumPoints
    For j = i + 1 To m_NumPoints
        Distances(i, j) = Sqr((m_PointX(j) - m_PointX(i)) ^ 2 + (m_PointY(j) - m_PointY(i)) ^ 2)
        Distances(j, i) = Distances(i, j)
    Next
Next i

' Calculate Permutations..
NumValues = m_NumPoints - 2
ReDim Used(1 To NumValues)
ReDim CurrentSolution(0 To Factorial(NumValues) - 1, 1 To NumValues)
CurSolNum = 0
EnumerateValues 1

MinimumLength = 1000000
' Calculate the sum of distances in every permutation
' always saving the lowest solution
For Route = 0 To Factorial(NumValues) - 1
    MySum = Distances(m_StartHandle, CurrentSolution(Route, 1))
    For PointNum = 1 To NumValues - 1
        MySum = MySum + Distances(CurrentSolution(Route, PointNum), CurrentSolution(Route, PointNum + 1))
    Next PointNum
    MySum = MySum + Distances(CurrentSolution(Route, NumValues), m_EndHandle)
    If MySum < MinimumLength Then
        MinimumLength = MySum
        ShortestPath = Route
    End If
Next Route

ReDim InShortest(1 To m_NumPoints, 1 To m_NumPoints)

' Find the vertices in the solution in order to color them different
Label1.Caption = "Length: " & vbCrLf & MinimumLength & vbCrLf & m_StartHandle
OldCurSol = m_StartHandle
For PointNum = 1 To NumValues
    CurSol = CurrentSolution(ShortestPath, PointNum)
    InShortest(CurSol, OldCurSol) = True
    InShortest(OldCurSol, CurSol) = True
    Label1.Caption = Label1.Caption & vbCrLf & CurSol
    OldCurSol = CurSol
Next
InShortest(CurSol, m_EndHandle) = True
InShortest(m_EndHandle, CurSol) = True

Label1.Caption = Label1.Caption & vbCrLf & m_EndHandle

DrawShortest = True
Refresh

'Stop
End Sub

Private Sub Form_Activate()

'Draw shortest path
cmdShortPath_Click

End Sub

' Create an initial diamond of 5.
Private Sub Form_Load()
    
Randomize

EdgeColor = vbBlack
ShortEdgeColor = vbRed
VertexColor = vbWhite
VertexBorderColor = vbBlack
StartVertexColor = vbGreen
EndVertexColor = vbRed
TextColor = vbBlue

DrawShortest = False

m_StartHandle = 4
m_EndHandle = 5

Vertex_Change (5)

End Sub
' See if we are over a grab handle.
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
Dim dx As Single
Dim dy As Single

If Button = 2 Then
    ' You can use the right mouse button to place/remove crosses
    DelCross = False
    For t = 1 To NumCross
        If Sqr((Cross(t).x - x) ^ 2 + (Cross(t).y - y) ^ 2) < 5 Then
            DelCross = True
            Exit For
        End If
    Next
    If DelCross Then
        Cross(t).x = -10
        Cross(t).y = -10
        Refresh
        If ReelNumCross = 1 Then
            ReDim Cross(1 To 1)
            NumCross = 0
        End If
        ReelNumCross = ReelNumCross - 1
    Else
        NumCross = NumCross + 1
        ReelNumCross = ReelNumCross + 1
        ReDim Preserve Cross(1 To NumCross) As POINTAPI
        Cross(NumCross).x = x
        Cross(NumCross).y = y
        For t = 1 To NumCross
            Me.Line (Cross(t).x - 2, Cross(t).y - 2)-(Cross(t).x + 3, Cross(t).y + 3)
            Me.Line (Cross(t).x - 2, Cross(t).y + 2)-(Cross(t).x + 3, Cross(t).y - 3)
        Next
    End If
Else
    For i = 1 To m_NumPoints
        If Abs(m_PointX(i) - x) < HANDLE_HALF_WIDTH And _
           Abs(m_PointY(i) - y) < HANDLE_HALF_WIDTH _
        Then
            ' We are over this grab handle.
            ' Start dragging.
            m_DraggingHandle = i
            ' If the number of vertexes are less than 9 then draw shortest path
            If m_NumPoints < 9 Then cmdShortPath_Click Else Refresh
            Exit For
        End If
    Next i
End If
End Sub

' Move the drag handle.
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Do nothing if we are not dragging.
    If m_DraggingHandle = 0 Then Exit Sub
    
    ' Move the handle.
    m_PointX(m_DraggingHandle) = x
    m_PointY(m_DraggingHandle) = y
    
    'The shortest may have changes
    DrawShortest = False
    
    ' Redraw.
    ' If AutoDraw is enabled then draw shortest path
    If chkAutoDraw.Value = vbChecked Then cmdShortPath_Click Else Refresh
End Sub


' Stop dragging.
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_DraggingHandle = 0

End Sub


Private Sub Form_Paint()
Dim i As Integer, j As Integer, Dist As Single

    Cls
    If m_NumPoints < 1 Then Exit Sub

    ' Start at the last point.
    CurrentX = m_PointX(m_NumPoints)
    CurrentY = m_PointY(m_NumPoints)

    ' Connect the points.
    ForeColor = EdgeColor
    For i = 1 To m_NumPoints
        For j = i + 1 To m_NumPoints
            If DrawShortest Then
                If InShortest(i, j) = True Then ForeColor = ShortEdgeColor Else ForeColor = EdgeColor
            End If
            If chkEdges.Value = vbChecked Or ForeColor = ShortEdgeColor Then
                Line (m_PointX(j), m_PointY(j))-(m_PointX(i), m_PointY(i))
                CurrentX = (m_PointX(j) + m_PointX(i)) / 2
                CurrentY = (m_PointY(i) + m_PointY(j)) / 2
                If chkDist.Value = vbChecked Then
                    ForeColor = TextColor
                    Dist = Sqr((m_PointX(j) - m_PointX(i)) ^ 2 + (m_PointY(j) - m_PointY(i)) ^ 2) / 30 'Conv from pixel to cm
                    Print Round(Dist, 1)
                    ForeColor = EdgeColor
                End If
            End If
        Next
    Next i

    ' Draw grab handles as white squares with
    ' black edges.
    FillColor = VertexColor
    FillStyle = vbFSSolid
    ForeColor = VertexBorderColor
    For i = 1 To m_NumPoints
        If i > m_NumPoints - 2 Then '= m_StartHandle Or i = m_EndHandle Then
            If i = m_NumPoints - 1 Then ' m_StartHandle Then
                FillColor = StartVertexColor
                m_StartHandle = i
            Else
                FillColor = EndVertexColor
                m_EndHandle = i
            End If
            Line (m_PointX(i) - HANDLE_HALF_WIDTH, m_PointY(i) - HANDLE_HALF_WIDTH)-Step(HANDLE_WIDTH, HANDLE_WIDTH), , B
            Print i
            FillColor = VertexColor
        Else
            Line (m_PointX(i) - HANDLE_HALF_WIDTH, m_PointY(i) - HANDLE_HALF_WIDTH)-Step(HANDLE_WIDTH, HANDLE_WIDTH), , B
            Print i
        End If
    Next i
    
    'Draw marks
    For t = 1 To NumCross
        Me.Line (Cross(t).x - 2, Cross(t).y - 2)-(Cross(t).x + 3, Cross(t).y + 3)
        Me.Line (Cross(t).x - 2, Cross(t).y + 2)-(Cross(t).x + 3, Cross(t).y - 3)
    Next
    
End Sub


Private Sub lblNumber_Click()

End Sub

Private Sub txtVertexes_Change()
If Val(txtVertexes.Text) > 1 Then Vertex_Change (Val(txtVertexes.Text))

End Sub

Private Sub Vertex_Change(NumPoints As Integer)
Dim VertexNum As Integer

m_NumPoints = NumPoints

    'Make new caption   (Messy calculations I know!!!)
Select Case m_NumPoints
Case Is < 10
    EstTid = " <1 sec"
Case 10
    EstTid = " ~2 sec"
Case 11
    EstTid = " ~30 sec"
Case 12
    EstTid = " ~300 sec"
Case 13
    EstTid = " ~" & Round((5 / Factorial(12)) * Factorial(m_NumPoints), 2) & " min"
Case 14
    EstTid = " Unwise! : ~" & Round((5 / Factorial(12)) / 60 * Factorial(m_NumPoints), 2) & " hours"
Case 15
    EstTid = " Unwise! : ~" & Round((5 / Factorial(12)) / 60 / 24 * Factorial(m_NumPoints), 2) & " days"
Case 16
    EstTid = " Unwise! : ~" & Round((5 / Factorial(12)) / 60 / 24 * Factorial(m_NumPoints), 2) & " days"
Case 17, 18
    EstTid = " Unwise! : ~" & Round((5 / Factorial(12)) / 60 / 24 / 365 * Factorial(m_NumPoints), 2) & " years"
Case 19, 20, 21
    EstTid = " Unwise! : ~" & Int((5 / Factorial(12)) / 60 / 24 / 365 * Factorial(m_NumPoints)) & " years"
Case Is > 100
    EstTid = " Never!"
Case Is > 21
    mycalc1 = (5 / Factorial(12)) / 60 / 24 / 365 * Factorial(m_NumPoints)
    mycalc = Int(Log(mycalc1) / Log(10))
    EstTid = "Unwise! : ~" & Round(mycalc1 / 10 ^ mycalc, 2) & " * 10^" & mycalc & " years"
End Select
Me.Caption = "Shortest Path Calculator - Points: " & m_NumPoints & ", Est. Calc. Time: " & EstTid
    
    ' Make room.

If m_NumPoints > 1 Then
    ReDim Preserve m_PointX(1 To m_NumPoints)
    ReDim Preserve m_PointY(1 To m_NumPoints)

    ' Set initial points.
    For VertexNum = 1 To m_NumPoints
        m_PointX(VertexNum) = Cos(TwoPi / m_NumPoints * VertexNum) * 200 + 300
        m_PointY(VertexNum) = Sin(TwoPi / m_NumPoints * VertexNum) * 200 + 240
    Next

    'No longer valid
    DrawShortest = False

    ' Draw.
    Refresh
End If

End Sub

Private Sub txtVertexes_Click()
txtVertexes.SelStart = 0
txtVertexes.SelLength = Len(txtVertexes.Text)

End Sub

Private Sub EnumerateValues(ByVal index As Integer)

'************************************'
'*  EnumerateValues is written by   *'
'*        www.vb-helper.com         *'
'************************************'

Dim result As String
Dim i As Integer, t As Long, stopnum As Long

    ' See if there are any values left to try.
    If index > NumValues Then
        ' All values are used.
        ' Get a string for the current solution.
        For i = 1 To NumValues
            result = result & Format$(CurrentSolution(CurSolNum, i)) & " "
        Next i

        ' Add the current solution to the list.
        'txtResults.Text = txtResults.Text & result & vbCrLf
        'OldBegin = CurrentSolution(CurSolNum, 0)
        CurSolNum = CurSolNum + 1
        
        Exit Sub
    End If

    ' Examine each value.
    For i = 1 To NumValues
        ' See if this value has been used yet.
        If Not Used(i) Then
            ' It is unused. Try using it.
            Used(i) = True
            stopnum = CurSolNum + Factorial(NumValues - index) '/ Factorial(index)
            If stopnum > Factorial(NumValues) - 1 Then stopnum = Factorial(NumValues) - 1
            For t = CurSolNum To stopnum
                CurrentSolution(t, index) = i
            Next

            EnumerateValues index + 1
            
            Used(i) = False
        End If
    Next i
End Sub

Private Function Factorial(ByVal num As Double) As Double
    If num <= 2 Then
        Factorial = num
    Else
        Factorial = num * Factorial(num - 1)
    End If
End Function








