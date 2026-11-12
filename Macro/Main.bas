Attribute VB_Name = "Main"
Public rinc As Integer, cinc As Integer
Dim r() As Integer, c() As Integer

Sub Start()

Range("B2:S26").Interior.Color = vbBlack
ReDim r(2)
ReDim c(2)
r(0) = 20: r(1) = 21: r(2) = 22
c(0) = 10: c(1) = 10: c(2) = 10
rinc = 0: cinc = 0
bindKeys
AddApple
ShowSnake
StartTimer
End Sub

Sub ShowSnake()
Dim i As Integer
For i = UBound(r) To 1 Step -1
  Cells(r(i), c(i)).Interior.Color = vbGreen
Next i
Cells(r(0), c(0)).Interior.Color = vbRed
End Sub

Sub MoveSnake()
Dim i As Integer
Dim tail As Integer
If rinc <> 0 Or cinc <> 0 Then
  tail = UBound(r)
  Cells(r(tail), c(tail)).Interior.Color = vbBlack
  
  'reposition the body of the snake
  For i = tail To 1 Step -1
    r(i) = r(i - 1)
    c(i) = c(i - 1)
  Next i
  
  'move the head of the snake
  r(0) = r(0) + rinc
  c(0) = c(0) + cinc
  
  'Check the movement of the snake
  If Cells(r(0), c(0)).Interior.Color = vbYellow Then
   'got the apple
   Apples = Apples + 1
   ReDim Preserve r(UBound(r) + 1)
   ReDim Preserve c(UBound(c) + 1)
   r(UBound(r)) = r(UBound(r) - 1)
   c(UBound(c)) = c(UBound(c) - 1)
   AddApple
 ElseIf Cells(r(0), c(0)).Interior.Color <> vbBlack Then
   StopTimer
   MsgBox "Game Over"
   Exit Sub
 End If
 
  ShowSnake
 End If
 
End Sub

Sub AddApple()
Randomize
Do
  arow = Int(Rnd * 25) + 2
  acol = Int(Rnd * 18) + 2
Loop While Cells(arow, acol).Interior.Color <> vbBlack

Cells(arow, acol).Interior.Color = vbYellow

End Sub

Sub ResumeGame()
    If GameRUnning = False Then
       GameRUnning = True
       bindKeys
       StartTimer
    End If
    
End Sub


