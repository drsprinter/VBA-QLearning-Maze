Sub QlearnRunning()
'Copyright 2017 Keisuke Hata
' QlearnRunning Macro
'
Dim wall, episodeNum, tate, yoko, randNum, solveFlag, stepCnt, score, stepFlag, randNumE, actionControlTate, actionControlYoko, Qdirection, errorDirection, Qstep As Integer
Dim color As Long
Dim alpha, gamma, price As Double
Dim Q(4), maxQ As Double
Dim multiFlag As Boolean
Dim multiCnt, multiArray(4), loopCnt As Integer


alpha = 1
gamma = 0.9
episodeNum = 0
multiFlag = False
multiCnt = 0


Do While episodeNum < 50


Qstep = 0
stepFlag = 0
tate = 2
yoko = 2
score = 0
solveFlag = 0
price = 0
wall = 0


Do While solveFlag = 0

randNumE = Int(Rnd * 99)
If randNumE < 0 Then

stepFlag = 0
loopCnt = 0
Do While stepFlag = 0
If loopCnt > 4 Then
Exit Do
End If
loopCnt = loopCnt + 1
randNum = Int(Rnd * 4)
actionControlTate = 0
actionControlYoko = 0
Do While actionControlTate = 0 And actionControlYoko = 0
If randNum = 0 Then
actionControlTate = 1
If tate < 12 Then
tate = tate + actionControlTate
End If

ElseIf randNum = 1 Then
actionControlTate = -1
If tate > 2 Then
tate = tate + actionControlTate




End If

ElseIf randNum = 2 Then
actionControlYoko = 1
If yoko < 14 Then
yoko = yoko + actionControlYoko



End If



ElseIf randNum = 3 Then
actionControlYoko = -1
If yoko > 2 Then
yoko = yoko + actionControlYoko

End If

End If

Loop

If Cells(tate, yoko).Interior.color <> RGB(31, 73, 125) Then





stepFlag = 1

Else
tate = tate - actionControlTate
yoko = yoko - actionControlYoko
End If

Loop 'step judge loop

If tate = 12 And yoko = 14 Then
price = 100
Else
price = 0
End If

Sheet2.Cells(3 * tate, 3 * yoko) = 1

Q(0) = Sheet2.Cells(3 * tate + 1, 3 * yoko)
Q(1) = Sheet2.Cells(3 * tate - 1, 3 * yoko)
Q(2) = Sheet2.Cells(3 * tate, 3 * yoko + 1)
Q(3) = Sheet2.Cells(3 * tate, 3 * yoko - 1)

maxQ = Q(0)

For i = 1 To 3
If maxQ < Q(i) Then
maxQ = Q(i)
End If
Next

If randNum = 0 Then

Q(0) = Sheet2.Cells(3 * (tate - actionControlTate) + 1, 3 * (yoko - actionControlYoko))
Q(0) = Q(0) + alpha * (price + gamma * maxQ - Q(0))
Sheet2.Cells(3 * (tate - actionControlTate) + 1, 3 * (yoko - actionControlYoko)) = Q(0)


ElseIf randNum = 1 Then

Q(1) = Sheet2.Cells(3 * (tate - actionControlTate) - 1, 3 * (yoko - actionControlYoko))
Q(1) = Q(1) + alpha * (price + gamma * maxQ - Q(1))
Sheet2.Cells(3 * (tate - actionControlTate) - 1, 3 * (yoko - actionControlYoko)) = Q(1)

ElseIf randNum = 2 Then
Q(2) = Sheet2.Cells(3 * (tate - actionControlTate), 3 * (yoko - actionControlYoko) + 1)
Q(2) = Q(2) + alpha * (price + gamma * maxQ - Q(2))
Sheet2.Cells(3 * (tate - actionControlTate), 3 * (yoko - actionControlYoko) + 1) = Q(2)

ElseIf randNum = 3 Then
Q(3) = Sheet2.Cells(3 * (tate - actionControlTate), 3 * (yoko - actionControlYoko) - 1)
Q(3) = Q(3) + alpha * (price + gamma * maxQ - Q(3))
Sheet2.Cells(3 * (tate - actionControlTate), 3 * (yoko - actionControlYoko) - 1) = Q(3)


End If


If Cells(tate, yoko).Interior.color <> RGB(255, 0, 0) Then
Cells(tate, yoko).Interior.color = RGB(0, 255, 0)
End If

If tate = 12 And yoko = 14 Then
solveFlag = 1
End If
stepCnt = stepCnt + 1
Cells(3, 23) = stepCnt
Cells(1, 23) = randNumE

Else 'randNumE
Q(0) = Sheet2.Cells(3 * tate + 1, 3 * yoko)
Q(1) = Sheet2.Cells(3 * tate - 1, 3 * yoko)
Q(2) = Sheet2.Cells(3 * tate, 3 * yoko + 1)
Q(3) = Sheet2.Cells(3 * tate, 3 * yoko - 1)




stepFlag = 0
loopCnt = 0
Do While stepFlag = 0
If loopCnt > 4 Then
Exit Do
End If
loopCnt = loopCnt + 1
randNum = Int(Rnd * 4)
actionControlTate = 0
actionControlYoko = 0
maxQ = 0

Qdirection = 0
multiCnt = 0
multiFlag = False
For i = 0 To 3
If maxQ < Q(i) Then
maxQ = Q(i)
Qdirection = i

multiFlag = False
multiArray(0) = i
ElseIf maxQ = Q(i) And i <> 0 Then
multiFlag = True
multiArray(multiCnt + 1) = i
multiCnt = multiCnt + 1

ElseIf maxQ = Q(i) And i = 0 Then
multiArray(0) = 0
End If
Next

If multiFlag = True Then
randNum = Int(Rnd * (multiCnt + 1))
maxQ = Q(multiArray(randNum))
Qdirection = multiArray(randNum)
End If


actionControlTate = 0
actionControlYoko = 0
If Qdirection = 0 Then
actionControlTate = 1
If tate < 12 Then
tate = tate + actionControlTate
End If

ElseIf Qdirection = 1 Then
actionControlTate = -1
If tate > 2 Then
tate = tate + actionControlTate




End If

ElseIf Qdirection = 2 Then
actionControlYoko = 1
If yoko < 14 Then
yoko = yoko + actionControlYoko



End If



ElseIf Qdirection = 3 Then
actionControlYoko = -1
If yoko > 2 Then
yoko = yoko + actionControlYoko



End If
End If

If Cells(tate, yoko).Interior.color <> RGB(31, 73, 125) Then




Sheet2.Cells(3 * tate, 3 * yoko) = 1
stepFlag = 1

Else
tate = tate - actionControlTate
yoko = yoko - actionControlYoko
errorDirection = Qdirection
End If

Loop 'step judge loop

If tate = 12 And yoko = 14 Then
price = 100
Else
price = 0
End If

Sheet2.Cells(3 * tate, 3 * yoko) = 1

Q(0) = Sheet2.Cells(3 * tate + 1, 3 * yoko)
Q(1) = Sheet2.Cells(3 * tate - 1, 3 * yoko)
Q(2) = Sheet2.Cells(3 * tate, 3 * yoko + 1)
Q(3) = Sheet2.Cells(3 * tate, 3 * yoko - 1)

maxQ = Q(0)
Qdirection = 0
For i = 1 To 3
If maxQ < Q(i) Then
maxQ = Q(i)
'Qdirection = i
End If
Next


'Sheet2.Cells(3 * tate, 3 * yoko) = maxQ

'If Qdirection = 0 Then
'
'Q(0) = Sheet2.Cells(3 * (tate - actionControlTate) + 1, 3 * (yoko - actionControlYoko))
'Q(0) = Q(0) + alpha * (price + gamma * maxQ - Q(0))
'Sheet2.Cells(3 * (tate - actionControlTate) + 1, 3 * (yoko - actionControlYoko)) = Q(0)
'
'
'ElseIf Qdirection = 1 Then
'
'Q(1) = Sheet2.Cells(3 * (tate - actionControlTate) - 1, 3 * (yoko - actionControlYoko))
'Q(1) = Q(1) + alpha * (price + gamma * maxQ - Q(1))
'Sheet2.Cells(3 * (tate - actionControlTate) - 1, 3 * (yoko - actionControlYoko)) = Q(1)
'
'ElseIf Qdirection = 2 Then
'Q(2) = Sheet2.Cells(3 * (tate - actionControlTate), 3 * (yoko - actionControlYoko) + 1)
'Q(2) = Q(2) + alpha * (price + gamma * maxQ - Q(2))
'Sheet2.Cells(3 * (tate - actionControlTate), 3 * (yoko - actionControlYoko) + 1) = Q(2)
'
'ElseIf Qdirection = 3 Then
'Q(3) = Sheet2.Cells(3 * (tate - actionControlTate), 3 * (yoko - actionControlYoko) - 1)
'Q(3) = Q(3) + alpha * (price + gamma * maxQ - Q(3))
'Sheet2.Cells(3 * (tate - actionControlTate), 3 * (yoko - actionControlYoko) - 1) = Q(3)
'
'
'End If


If Cells(tate, yoko).Interior.color <> RGB(255, 0, 0) Then
Cells(tate, yoko).Interior.color = RGB(0, 255, 0)
End If

If tate = 12 And yoko = 14 Then
solveFlag = 1
End If
stepCnt = stepCnt + 1
Cells(3, 23) = stepCnt
Qstep = Qstep + 1
Cells(4, 23) = Qstep
Cells(1, 23) = randNumE

End If 'randNumE



Application.Wait [Now() + "0:00:00.01"]
Loop



episodeNum = episodeNum + 1
If episodeNum < 200 Then
ClearRout
End If

Cells(episodeNum, 28) = stepCnt
stepCnt = 0
Loop



Cells(2, 23) = score




'
End Sub
