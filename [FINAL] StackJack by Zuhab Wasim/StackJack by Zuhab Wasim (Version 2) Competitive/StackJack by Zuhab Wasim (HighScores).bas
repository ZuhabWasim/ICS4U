Attribute VB_Name = "Module3"
Option Explicit

'Highscore record type
Type HighScoreRec
    HName As String * 25
    HScore As Integer
    HTime As Single
End Type

'All variables and constants relating to the highscores
Global Const HIGHSCORE_FILE = "HighScores.rec"
Global Const HIGHSCORE_MAX = 10
Global HighScoreLength As Integer
Global HighScores(1 To HIGHSCORE_MAX) As HighScoreRec

'NOTE: There is no need to kill the filename for any of these procedures
'      because the file ensures its existence as it is created during form load if non-existent

'Creates the record file for highscores if it does not already exist
Public Sub CreateFile()

    Dim K As Integer
    Dim PlaceHolder As HighScoreRec
    
    'Placeholder for each record element
    With PlaceHolder
        .HName = "Anonymous"
        .HScore = 0
        .HTime = 0
    End With
    
    Open App.Path & "\" & HIGHSCORE_FILE For Random As #1 Len = HighScoreLength
    
    For K = 1 To HIGHSCORE_MAX
        Put #1, K, PlaceHolder
    Next K
    
    Close #1

End Sub

'Updates the file with the new highscores
Public Sub SaveFile(ByRef HighScores() As HighScoreRec, _
                    ByVal FileName As String, _
                    ByVal HighScoreLen As Integer)
    
    Dim K As Integer
    
    Open App.Path & "\" & HIGHSCORE_FILE For Random As #1 Len = HighScoreLen
    For K = 1 To HIGHSCORE_MAX
        Put #1, K, HighScores(K)
    Next K
    
    Close #1

End Sub

'Retrieves all the highscores from the highscore record file
'and puts it into the HighScores record
Public Sub GetHighScores(ByRef HighScores() As HighScoreRec, _
                         ByVal FileName As String, _
                         ByVal HighScoreLen As Integer)
                                         
    Dim K As Integer
    
    Open App.Path & "\" & HIGHSCORE_FILE For Random As #1 Len = HighScoreLen
    For K = 1 To HIGHSCORE_MAX
        Get #1, K, HighScores(K)
    Next K
    
    Close #1
    
End Sub

'Determines the position of the score if it reaches the leaderboard
Public Function NewHighScorePos(ByRef HighScores() As HighScoreRec, _
                                ByVal CurrentScore As Integer, _
                                ByVal CurrentTime As Single) As Integer
    
    Dim K As Integer
    
    'Initially sets the position to -1 (assuming it not a highscore)
    NewHighScorePos = -1
    
    For K = 1 To HIGHSCORE_MAX
        If CurrentScore > HighScores(K).HScore Then
            NewHighScorePos = K
            Exit For
        ElseIf CurrentScore = HighScores(K).HScore Then
            'Uses time as a tie breaker for scores
            If CurrentTime < HighScores(K).HTime Then
                NewHighScorePos = K
                Exit For
            End If
        End If
    Next K
    
End Function

'Updates the highscore list with the new highscores
Public Sub ChangeHighScores(ByRef HighScores() As HighScoreRec, _
                            ByVal CurrentScore As Integer, _
                            ByVal CurrentTime As Single, _
                            ByVal NickName As String, _
                            ByVal CurrentPos As Integer)

    Dim K As Integer
    Dim TempName As String
    Dim TempScore As Integer
    Dim TempTime As Single
    
    'Swaps all elements from the lowest to the newly achieved highscore
    For K = HIGHSCORE_MAX To CurrentPos + 1 Step -1
        SwapName HighScores(K).HName, HighScores(K - 1).HName
        SwapInteger HighScores(K).HScore, HighScores(K - 1).HScore
        SwapSingle HighScores(K).HTime, HighScores(K - 1).HTime
    Next K
    
    'Finally assigns the newly achieved highscore to its given position
    HighScores(K).HName = NickName
    HighScores(K).HScore = CurrentScore
    HighScores(K).HTime = CurrentTime
    
End Sub

'Swaps places of two string values
Public Sub SwapName(ByRef Name1 As String, _
                    ByRef Name2 As String)
    
    Dim TempName As String
    
    TempName = Name1
    Name1 = Name2
    Name2 = TempName
    
End Sub

'Swaps places of two integer values
Public Sub SwapInteger(ByRef Num1 As Integer, _
                       ByRef Num2 As Integer)
    
    Dim TempNum As Integer
    
    TempNum = Num1
    Num1 = Num2
    Num2 = TempNum
    
End Sub

'Swaps places of two single values
Public Sub SwapSingle(ByRef Num1 As Single, _
                      ByRef Num2 As Single)
    
    Dim TempNum As Single
    
    TempNum = Num1
    Num1 = Num2
    Num2 = TempNum
    
End Sub


