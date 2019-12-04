Sub FindWinners()

Dim First As Long
Dim Second As Long
Dim Third As Long
Dim Wild1() As Variant
Dim Player() As Variant

'Initialize the variables with the winning numbers
First = 3957481
Second = 5865187
Third = 2817729
'Initialize the Wild1 array with wildcard winning numbers (any of which will win a runner up prize)
Wild1 = Array(2275339, 5868182, 1841402)

'Loop through all the rows and match against the winning numbers for first,second and third
For i = 2 To 1001
    'Place the firstname (col1), LastName (col2) and LottoNumber (col3) into an two dimensional array Player()
    Player = Range("A" & i & ":C" & i).Value
    
    'Compare the LottoNumber which is the array value 1,3 from the Player array against the first winning number
    If Player(1, 3) = First Then
        'If there is a match write the array contents to the appropriate cells in the spreadsheet
        Range("F2:H2").Value = Player
        'Notify the winner of first place with a message box
        MsgBox ("Congratulations " & Player(1, 1) & " " & Player(1, 2) & " you won First Place!")
        'If there is a match write the array contents to the appropriate cells in the spreadsheet
    ElseIf Player(1, 3) = Second Then
        Range("F3:H3").Value = Player
    ElseIf Player(1, 3) = Third Then
        Range("F4:H4").Value = Player
    End If
Next i

'Go through the list of players and their LottoNumber to see if there is a match to any of the wildcard winning numbers in the Wild1 array
For i = 2 To 1001
    Player = Range("A" & i & ":C" & i).Value
    'Compare the LottoNumber of a Player to see if it is contained in the Wild1 array using a filter function to return the matching value in the Wild1 array and then determining if the largest index is greater than -1 (if a value is returned it will have at least a value of 0)
    If (UBound(Filter(Wild1, Player(1, 3))) > -1) Then
        Range("F5:H5").Value = Player
        Exit For
    End If
Next i


End Sub