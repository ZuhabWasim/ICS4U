Attribute VB_Name = "Module1"
Option Explicit
'General code for the program

Public Sub Display(ByRef Cards() As String, ByVal MAX As Integer)
    
    'Local variables declared
    Dim K As Integer
    Dim NewSt As String
    
    'Iterates through the array and displays all credit cards
    For K = 1 To MAX
        'Checks to see if the credit card number is greater than 20 characters and cuts it of if so
        If Len(Cards(K)) > 25 Then
            NewSt = Left$(Cards(K), 25) & "..."
        Else
            NewSt = Cards(K)
        End If
        'Displays the credit cards on the picture box picData
        frmMain.picData.Print " " & Format$(K, "@@") & ". " & NewSt
    Next K

End Sub

Public Sub ReadFile(ByVal FName As String, ByRef CC() As String, ByRef TCards)

    'Local variables declared
    Dim K As Integer
    
    'Variables initialized
    K = 0
    
    'Opens the given file name to retrieve input
    Open FName For Input As #1
    
    'Iterates through the sequential file, storing every entry into an array
    Do While Not EOF(1)
        K = K + 1
        Input #1, CC(K)
    Loop
    
    'Closes the file
    Close #1
    
    'Assigns the total amount of entries in the variable TotalCards to be later used for efficiency purposes
    TCards = K
    
End Sub

Public Sub ValidateCards(ByRef CC() As String, ByVal TotalCards As Integer, ByRef ACount As Integer, ByRef MCount As Integer, ByRef VCount As Integer, ByRef NValid As Integer)
    
    'Local variables declared
    Dim K As Integer
    Dim Digits As Integer
    
    'Counted loop goes through every card in the array
    For K = 1 To TotalCards
    
        'Retrieves the first 2 numbers in the string and converts it into a numeric value
        Digits = Val(Left$(CC(K), 2))
        
        'Checks to see if the card is valid
        If ValidCard(CC(K)) Then
            'If the card has the length of 16 characters, we know it is either Mastercard or Visa
            If Len(CC(K)) = 16 Then
                'If the first two digits have the value between 51 and 55, it is a mastercard
                If Digits >= 51 And Digits <= 55 Then
                    MCount = MCount + 1
                'Otherwise if the card starts with the digit 4 then we know it is a visa
                ElseIf Mid$(Str$(Digits), 2, 1) = "4" Then
                    VCount = VCount + 1
                End If
            'If the card has the length of 15 characters, and the first 2 digits are 34 or 37, it is a amex
            ElseIf Len(CC(K)) = 15 And (Digits = 34 Or Digits = 37) Then
                ACount = ACount + 1
            End If
        Else
            'If the card is not valid, increment the not valid variable
            NValid = NValid + 1
        End If
        
    Next K
    
End Sub


Public Function ValidCard(ByVal Card As String) As Boolean
    
    'Local variables declared
    Dim K As Integer
    Dim Start As Integer
    Dim NewCard As String
    Dim Product As Integer
    Dim Sum As Integer
    
    'Variables initialized
    Sum = 0
    NewCard = ""
    
    'Checks to see where the loop should start, even is at 1, odd is at 2
    If Len(Card) Mod 2 = 0 Then
        Start = 1
    Else
        'If it is odd, we must account for the first digit, assigning it to the new card variable
        'and adding it to the sum
        Start = 2
        NewCard = Left$(Card, 1)
        Sum = Val(Left$(Card, 1))
    End If
    
    'Loops through every other characters in the card
    For K = Start To Len(Card) Step 2
        'Doubles the current card digit
        Product = Val(Mid$(Card, K, 1)) * 2
        'If the card is greater than 10, we must add the digits, otherwise simply put it into the newcard string
        If Product >= 10 Then
            Product = Val(Mid$(Str$(Product), 2, 1)) + Val(Mid$(Str$(Product), 3, 1))
        End If
        'Accounts for each character skipped in the string
        'The product and skipped character are added for the sum, and are concatenated for the NewCard variable
        NewCard = NewCard & Trim$(Str$(Product)) & Mid$(Card, K + 1, 1)
        Sum = Sum + Product + Val(Mid$(Card, K + 1, 1))
    Next K

    'Checks to see if that new sum is divisible by 10 and is not 0, it is valid if it is
    If Sum Mod 10 = 0 And Sum <> 0 Then
        ValidCard = True
    Else
        ValidCard = False
    End If
    
End Function
