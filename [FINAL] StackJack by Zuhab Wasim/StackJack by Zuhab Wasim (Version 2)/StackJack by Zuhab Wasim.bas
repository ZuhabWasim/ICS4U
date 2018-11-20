Attribute VB_Name = "Module2"
Option Explicit

'Gameover checker variable
Global GameOver As Boolean
Global GameOverMessage As String

'Pause variable
Global GamePausable As Boolean

'Initializes all variables and arrays used by the game
Public Sub Initialize(ByRef Cards() As Long, _
                      ByVal MAX_CARDS As Integer, _
                      ByRef CardCounter As Integer, _
                      ByRef Stacks() As Integer, _
                      ByRef StackCount() As Integer, _
                      ByRef StackPoints() As Integer, _
                      ByRef Points As Integer, _
                      ByRef ScoreMultiplier As Integer, _
                      ByRef HighValueAce() As Boolean, _
                      ByVal MAX_STACKS As Integer, _
                      ByRef TimeCount As Single)

    Dim K As Integer
    
    'Initializes values
    CardCounter = 0
    Points = 0
    TimeCount = 0
    
    'Score multiplier is 1 because it is multiplied with the score
    ScoreMultiplier = 1
    
    GamePausable = False
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
                       ByVal Points As Integer, _
                       ByRef lblCardsLeft As Control, _
                       ByVal MAX_CARDS As Integer)
                       
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
    
    'Changes the colour of the cards remaining to show severity
    Select Case CardCounter
        'No more cards is black
        Case MAX_CARDS
            lblCardsLeft.ForeColor = QBColor(0)
        'The first 2 sets of 13 cards are green
        Case 1 To (MAX_CARDS / 2)
            lblCardsLeft.ForeColor = QBColor(2)
        'The last 26 are yellow
        Case (MAX_CARDS / 2 + 1) To (MAX_CARDS / (4 / 3))
            lblCardsLeft.ForeColor = QBColor(6)
        'The last 13 cards are in red
        Case (MAX_CARDS / (4 / 3) + 1) To MAX_CARDS - 1
            lblCardsLeft.ForeColor = QBColor(4)
    End Select
    
    lblCardsLeft.Caption = Str$(MAX_CARDS - CardCounter)
    'Refreshes the picture box to see the new card
    picCard.Refresh
    
End Sub

'Busts the stack or disables a stack for no longer use by the user
Public Sub BustStack(ByRef lblCount As Control, _
                     ByRef picStack As Variant, _
                     ByVal currentPicBox As Integer, _
                     ByVal MAXSTACKS As Integer)
    
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
        GameOverMessage = "You lose! You have busted all your stacks and lost any of the points you had. "
    End If

End Sub

'Resets all controls on the form
Public Sub ResetForm(ByRef lblCount As Variant, _
                     ByRef picStack As Variant, _
                     ByVal MAX_STACKS As Integer, _
                     ByRef lblPoints As Control, _
                     ByRef cmdDiscard As Control, _
                     ByRef lblTime As Control, _
                     ByRef lblMulti As Control, _
                     ByRef lblPointsGiven As Control)
    
    Dim K As Integer
    
    'Initializes all the properties of each button or label
    cmdDiscard.Enabled = True
    lblPoints.Caption = "0"
    lblMulti.Caption = "1x"
    lblTime.Caption = "0:00.0"
    lblPointsGiven.Caption = ""
    lblPointsGiven.ForeColor = QBColor(0)
    
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
                   ByVal PTime As Single, _
                   ByRef cmdDiscard As Control, _
                   ByRef picStack As Variant, _
                   ByVal MAX_STACKS As Integer, _
                   ByRef mnuNewGame As Menu, _
                   ByRef tmrTimer As Control)
    
    Dim GOMessage As String
    Dim GOType As Integer
    Dim GOTitle As String

    Dim K As Integer
        
    Dim NickName As String
    Dim ScorePos As Integer
    
    'Disables all controls
    tmrTimer.Enabled = False
    cmdDiscard.Enabled = False
    mnuNewGame.Enabled = True
    For K = 1 To MAX_STACKS
        picStack(K).Enabled = False
    Next K
    
    'The game cannot pause anymore because the game has ended
    GamePausable = False
    
    'Shows the user their final score
    GOTitle = "Game Over!"
    GOType = vbInformation + vbOKOnly
    GOMessage = "Game Over! " & GameOverMessage & vbCrLf & "Your final score is " & Points & " points."
    
    MsgBox GOMessage, GOType, GOTitle
    
    'Checks to see if the user has "won" no highscoring if not
    If VBA.Left$(GameOverMessage, 7) = "You win" Then
        'Gets the position of the score on the leaderboard
        ScorePos = NewHighScorePos(HighScores(), Points, PTime)
        'If it exists, asks the user for a name and puts it on the leaderboard
        If ScorePos <> -1 Then
            GOTitle = "New HighScore!"
            GOMessage = "You have achieved a new high score! Please enter your name below: "
            NickName = InputBox$(GOMessage, GOTitle)
            'Error checks to see if the length of the given name has a length more than 25
            If Len(NickName) = 0 Then
                NickName = "Anonymous"
            ElseIf Len(NickName) > 25 Then
                NickName = VBA.Left$(NickName, 22) & "..."
            End If
            'Updates the scores
            ChangeHighScores HighScores(), Points, PTime, NickName, ScorePos
            'Saves the scores to the record file
            SaveFile HighScores(), HIGHSCORE_FILE, HighScoreLength
        End If
    End If
        
End Sub

'Updates the points awarded to the form and variable
Public Sub UpdatePoints(ByRef lblPoints As Control, _
                        ByRef Points As Integer, _
                        ByVal PValue As Integer, _
                        ByRef lblPointsGiven As Control)
    
    'Adds the points
    Points = Points + PValue
    
    'Checks to see if the points make the score go below 0 and resets it back to 0
    If Points < 0 Then
        'Determines if the points have been deducted to 0 and were not at 0
        'if they were display the subtraction back down to 0 otherwise dont display anything
        If Points <> PValue Then
            lblPointsGiven.ForeColor = QBColor(4)
            lblPointsGiven.Caption = Str$(PValue - Points)
        Else
            lblPointsGiven.Caption = ""
        End If
        Points = 0
    Else
        'Determines what to output for the points given caption
        'Add: "+VALUE" in green colour
        'Deduct: "-VALUE" in red colour
        'No change: Display nothing
        If PValue > 0 Then
            lblPointsGiven.ForeColor = QBColor(2)
            lblPointsGiven.Caption = "+" & Trim$(Str$(PValue))
        ElseIf PValue = 0 Then
            lblPointsGiven.Caption = ""
        Else
            lblPointsGiven.ForeColor = QBColor(4)
            lblPointsGiven.Caption = Str$(PValue)
        End If
    End If
    
    'Update points
    lblPoints.Caption = VBA.Str$(Points)
    
End Sub

'Returns the amount of points given from the stack clear
'based on the number of consecutive clears
Public Function StackPoints(ByVal ClearPoints As Integer, _
                            ByRef LastClear As Integer, _
                            ByVal CurrentClear As Integer, _
                            ByRef Multiplier As Integer) As Integer
                                
    If CurrentClear - 1 = LastClear Then
        Multiplier = Multiplier + 1
    Else
        Multiplier = 1
    End If
    
    StackPoints = ClearPoints * Multiplier
        
End Function

'Converts time in seconds to the format HH:MM:SS.S or M:SS.S
Public Function ConvTime(ByVal Seconds As Single) As String
    
    Dim Hours As String
    Dim Minutes As String
    Dim NewSeconds As String
    
    'Formats so that hours are not shown if not needed
    'and minutes are only formatted with leading zeros if time has hours
    If Seconds < 3600 Then
        Hours = ""
        'Since each time value is now a single and in miliseconds,
        'Int(Num / N) is chosen over Num \ N to avoid rounding issues
        Minutes = Int((ModDouble(Seconds, 3600)) / 60) & ":"
    Else
        Hours = Int(Seconds / 3600#) & ":"
        Minutes = VBA.Format$(Int(ModDouble(Seconds, 3600#) / 60), "00") & ":"
    End If
    NewSeconds = VBA.Format$((ModDouble(Seconds, 60#)), "00.0")
    
    ConvTime = Hours & Minutes & NewSeconds
    
End Function

'Centers the form on the screen
Public Sub CenterForm(ByRef givenForm As Form)
    
    givenForm.Move (Screen.Width - givenForm.Width) / 2, (Screen.Height - givenForm.Height) / 2
    
End Sub

'Used to evaluate the remainder of two double parameters as the mod operator does not suffice
Public Function ModDouble(ByVal Div1 As Double, _
                          ByVal Div2 As Double) As Double
                          
    Dim Num As Double
    
    Num = Div1
    
    Do While Num >= 0
        Num = Num - Div2
    Loop
    
    Num = Num + Div2
    
    ModDouble = Num
    
End Function

