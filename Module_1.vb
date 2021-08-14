Option Base 1

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms as Long)
#End If

Public stopNow As Boolean
Const row = 30
Const col = 30
Const ttS = 30

Sub theGameOfLife()
    Dim myArr(row, col) As Long
    Dim tempArr(row, col) As Long
    Dim sh_This As Worksheet
    stopNow = False
    
    'for user input seed
    Set sh_This = ThisWorkbook.Sheets(1)
    sh_This.Range("AG4").Select
    For i = 1 To row: For j = 1 To col
        tempColor = sh_This.Cells(i, j).Interior.Color
        
        If tempColor = vbWhite Then myArr(i, j) = 0
        If tempColor = vbBlack Then myArr(i, j) = 1
    Next j: Next i
    

    For steps = 1 To 1000
        'perform logic for next step
        If stopNow = True Then Exit Sub
        
        For i = 1 To row: For j = 1 To col
            'find out value of surrounding cells minus our cell
            mySum = 0
            For k = -1 To 1: For L = -1 To 1
                tempRow = (i + k + row - 1) Mod row + 1
                tempCol = (j + L + col - 1) Mod col + 1
            
                mySum = mySum + myArr(tempRow, tempCol)
            Next L: Next k
            
            mySum = mySum - myArr(i, j)
            
            'Any live cell with fewer than two live neighbours dies, as if by underpopulation.
            If myArr(i, j) = 1 And mySum < 2 Then tempArr(i, j) = 0
            'Any live cell with two or three live neighbours lives on to the next generation.
            If myArr(i, j) = 1 And (mySum = 2 Or mySum = 3) Then tempArr(i, j) = 1
            'Any live cell with more than three live neighbours dies, as if by overpopulation.
            If myArr(i, j) = 1 And mySum > 3 Then tempArr(i, j) = 0
            'Any dead cell with exactly three live neighbours becomes a live cell, as if by reproduction.
            If myArr(i, j) = 0 And mySum = 3 Then tempArr(i, j) = 1
        Next j: Next i
        
        For i = 1 To row: For j = 1 To col
            myArr(i, j) = tempArr(i, j)
        Next j: Next i
        
        drawArea myArr
        Application.StatusBar = "Gen: " & steps
        DoEvents
        Sleep ttS
    Next steps
    
End Sub

Sub drawArea(myArr() As Long)
    Application.ScreenUpdating = False
    
    Dim rows As Long, cols As Long
    rows = UBound(myArr) - LBound(myArr) + 1
    cols = UBound(myArr, 2) - LBound(myArr, 2) + 1
    
    Dim sh_This As Worksheet
    Set sh_This = ThisWorkbook.Sheets(1)
    
    For i = 1 To rows: For j = 1 To cols
        If myArr(i, j) = 0 Then
            sh_This.Cells(i, j).Interior.Color = vbWhite
        Else
            sh_This.Cells(i, j).Interior.Color = vbBlack
        End If
    Next j: Next i
    
    Application.ScreenUpdating = True
End Sub

Sub randomSeed()
    Dim myArr(row, col) As Long
    
    Randomize
    For i = 1 To row: For j = 1 To col
        myArr(i, j) = (1 * Rnd())
    Next j: Next i
    
    drawArea myArr
End Sub

Sub clearBoard()
    Application.ScreenUpdating = False
    
    Dim sh_This As Worksheet
    Set sh_This = ThisWorkbook.Sheets(1)
    
    For i = 1 To row: For j = 1 To col
        sh_This.Cells(i, j).Interior.Color = vbWhite
    Next j: Next i
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Sub endGame()
    stopNow = True
End Sub
