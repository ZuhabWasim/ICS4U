Attribute VB_Name = "Module2"
Option Explicit

Public Sub Initialize(ByRef Cards() As Long, ByVal MAX_CARDS As Integer, ByRef CardCounter As Integer)

    Dim K As Integer
    
    CardCounter = 0
    For K = 1 To MAX_CARDS
        Cards(K) = K - 1
    Next K
    
End Sub

Public Sub ShuffleDeck(ByRef Cards() As Long, ByVal MAX_CARDS As Integer, ByVal MIN_CARDS As Integer)
    
    Dim K As Integer
    
    Dim Card1 As Integer
    Dim Card2 As Integer
    Dim TempCard As Integer
    
    For K = 1 To 1000
        Card1 = Int(Rnd() * (MAX_CARDS - MIN_CARDS + 1)) + MIN_CARDS
        Card2 = Int(Rnd() * (MAX_CARDS - MIN_CARDS + 1)) + MIN_CARDS
        
        TempCard = Cards(Card1)
        Cards(Card1) = Cards(Card2)
        Cards(Card2) = TempCard
    Next K

End Sub
