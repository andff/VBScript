' Define variables
Const EMPTY = 0 ' Empty cell marker
Const X = 1 ' Player 1 marker
Const O = 2 ' Player 2 marker
arrBoard = Array(EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY)
intCurrentPlayer = X ' Start with Player 1

' Function to display the board
Sub DisplayBoard()
    For i = 1 To 3
        MsgBox "|" & arrBoard(i) & "|" & arrBoard(i + 3) & "|" & arrBoard(i + 6) & "|"
    Next

    MsgBox "|" & arrBoard(7) & "|" & arrBoard(8) & "|" & arrBoard(9) & "|"
End Sub

' Function to check for a winning condition
Function CheckWinner()
    ' Check rows
    For i = 1 To 3
        If arrBoard(i) <> EMPTY And arrBoard(i) = arrBoard(i + 3) And arrBoard(i + 3) = arrBoard(i + 6) Then
            CheckWinner = arrBoard(i)
            Exit Function
        End If
    Next

    ' Check columns
    For i = 1 To 7 Step 3
        If arrBoard(i) <> EMPTY And arrBoard(i) = arrBoard(i + 1) And arrBoard(i + 1) = arrBoard(i + 2) Then
            CheckWinner = arrBoard(i)
            Exit Function
        End If
    Next

    ' Check diagonals
    If arrBoard(1) <> EMPTY And arrBoard(1) = arrBoard(5) And arrBoard(5) = arrBoard(9) Then
        CheckWinner = arrBoard(1)
        Exit Function
    End If

    If arrBoard(7) <> EMPTY And arrBoard(7) = arrBoard(5) And arrBoard(5) = arrBoard(3) Then
        CheckWinner = arrBoard(7)
        Exit Function
    End If

    ' Check for a tie
    If Not IsEmpty(arrBoard) Then
