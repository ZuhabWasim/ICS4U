Attribute VB_Name = "Module2"
Option Explicit

'Gameover checker variable
Global GameOver As Boolean
Dim GameOverMessage As String

'Initializes all variables and arrays used by the game
Public Sub Initialize(ByRef Cards() As Long, _
                      ByVal MAX_CARDS As Integer, _
                      ByRef CardCounter As Integer, _
                      ByRef Stacks() As Integer, _
                      ByRef StackCount() As Integer, _
                      ByRef StackPoints() As Integer, _
                      ByRef Points As Integer, _
                      ByRef HighValueAce() As Boolean, _
                      ByVal MAX_STACKS As Integer)

    Dim K As Integer
    
    CardCounter = 0
    Points = 0
    GameOver = False
    
    'Assigns the values of the cards to the deck array
    For K = 1 To MAX_CARDS
        Cards(K) = K - 1
    Next K
    
    For K = 1 To MAX_STACKS
        Stacks(K) = 0
        StackCount(K) = 0
        StackPoints(K) = 0
        HighValueAce(K) = False
    Next K
    
End Sub

'Randomizes the sequence of the deck array
Public Sub ShuffleDeck(ByRef Cards() As Long, _
                       ByVal MAX_CARDS As Integer, _
                       ByVal MIN_CARDS As Integer)
    
    Dim K As Integer
    
    Dim Card1 As Integer
    Dim Card2 As Integer
    Dim TempCard As Integer
    
    'Swaps two random elements in the array 1000 times
    For K = 1 To 1000
        Card1 = Int(Rnd() * (MAX_CARDS - MIN_CARDS + 1)) + MIN_CARDS
        Card2 = Int(Rnd() * (MAX_CARDS - MIN_CARDS + 1)) + MIN_CARDS
        
        TempCard = Cards(Card1)
        Cards(Card1) = Cards(Card2)
        Cards(Card2) = TempCard
    Next K

End Sub

'Determines the specified value for each card
Public Function AddValue(ByVal Points As Integer, _
                         ByVal Card As Integer, _
                         ByRef HighValueAce As Boolean) As Integer
    
    Dim CardVal As Integer
    
    'Converts the values into a number representing the number of the card
    'i.e. an Ace of Spades and Ace of Hearts both have the value of 1
    CardVal = ((Card - 1) \ 4) + 1
    If Card Mod 4 = 0 And Card <> 0 Then
        CardVal = CardVal + 1
    End If
    
    'Uses cases to determine the value
    Select Case CardVal
        'Any card from 2 to 9 just has the value of that card assigned to it
        Case 2 To 9
            AddValue = CardVal
        'Any face card and the 10 card are all values at 10
        Case 10 To 13
            AddValue = 10
        'Checks to see if its valid to value the ace at 11 points instead of 1
        'and does so if it can.
        Case Else
            If Points + 11 <= 21 Then
                AddValue = 11
                HighValueAce = True
            Else
                AddValue = 1
            End If
    End Select
            
End Function

'Draws a new card from the deck array and prints it to the picture box
Public Sub DrawNewCard(ByRef picCard As Control, _
                       ByRef Cards() As Long, _
                       ByRef CardCounter As Integer, _
                       ByVal Points As Integer)
                       
    Dim Ret As Long
    
    'Increments to the next card
    CardCounter = CardCounter + 1
    
    'If all cards are used up, simply displays the back face of the card
    'and it becomes game over
    If CardCounter < 52 Then
        Ret = cdtDraw(picCard.hDC, 0, 0, Cards(CardCounter), C_FACES, 0)
    Else
        Ret = cdtDraw(picCard.hDC, 0, 0, Castle, C_BACKS, 0)
        GameOver = True
        If Points = 0 Then
            GameOverMessage = "You lose! You do not have anymore cards to use and your points are 0. "
        Else
            GameOverMessage = "You win! You have used the entire deck! "
        End If
    End If
    
    'Refreshes the picture box to see the new card
    picCard.Refresh
    
End Sub

'Busts the stack or disables a stack for no longer use by the user
Public Sub BustStack(ByRef lblCount As Control, _
                     ByRef picStack As Variant, _
                     ByVal currentPicBox As Integer, _
                     ByVal MAXSTACKS As Integer, _
                     ByRef Points As Integer)
    
    Dim K As Integer
    Dim CounterBust As Integer
    
    K = 0
    CounterBust = 0
    
    'Disables the column and changes the colour of the label signifying a bust
    lblCount.ForeColor = QBColor(12)
    picStack(currentPicBox).Enabled = False
    
    'Checks to see if all 5 columns are bust, assigns GameOver to true if so
    Do While K < 5
        K = K + 1
        If Not picStack(K).Enabled Then
            CounterBust = CounterBust + 1
        End If
    Loop
    
    If CounterBust = 5 Then
        GameOver = True
        GameOverMessage = "You lose! You have busted all your stacks. "
        Points = 0
    End If

End Sub

'Resets all controls on the form
Public Sub ResetForm(ByRef lblCount As Variant, _
                     ByRef picStack As Variant, _
                     ByVal MAX_STACKS As Integer, _
                     ByRef lblPoints As Control, _
                     ByRef cmdDiscard As Control)
    
    Dim K As Integer
    
    cmdDiscard.Enabled = True
    lblPoints.Caption = "0"

    For K = 1 To MAX_STACKS
        lblCount(K).Caption = "0"
        lblCount(K).ForeColor = QBColor(13)
        picStack(K).Enabled = True
        picStack(K).Cls
        picStack(K).Refresh
    Next K
    
End Sub

'Disables all user interaction and displays an endgame message
Public Sub EndGame(ByVal Points As Integer, _
                   ByRef cmdDiscard As Control, _
                   ByRef picStack As Variant, _
                   ByVal MAX_STACKS As Integer, _
                   ByRef mnuNewGame As Menu)
    
    Dim GOMessage As String
    Dim GOType As Integer
    Dim GOTitle As String
    Dim K As Integer
    
    'Disables all controls
    cmdDiscard.Enabled = False
    mnuNewGame.Enabled = True
    For K = 1 To MAX_STACKS
        picStack(K).Enabled = False
    Next K
    
    'Shows the user their final score
    GOTitle = "Game Over!"
    GOType = vbInformation + vbOKOnly
    GOMessage = "Game Over! " & GameOverMessage & vbCrLf & "Your final score is " & Points & " points."
    
    MsgBox GOMessage, GOType, GOTitle
    
End Sub

'Updates the points awarded to the form and variable
Public Sub UpdatePoints(ByRef lblPoints As Control, _
                        ByRef Points As Integer, _
                        ByVal PValue As Integer)
    
    'Adds the points
    Points = Points + PValue
    
    'Checks to see if the points make the score go below 0
    'and assigns it to 0 if so
    If Points < 0 Then
        Points = 0
    End If
    
    lblPoints.Caption = Str$(Points)
    
End Sub
