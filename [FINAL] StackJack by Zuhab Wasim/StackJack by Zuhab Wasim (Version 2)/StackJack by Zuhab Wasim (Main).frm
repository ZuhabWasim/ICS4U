VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StackJack"
   ClientHeight    =   5910
   ClientLeft      =   3840
   ClientTop       =   3390
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StackJack by Zuhab Wasim (Main).frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   9405
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   480
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   5
      Left            =   8160
      ScaleHeight     =   5505
      ScaleWidth      =   1065
      TabIndex        =   18
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdDiscard 
      Caption         =   "Discard Card"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   2280
      Width           =   1695
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   4
      Left            =   6720
      ScaleHeight     =   5505
      ScaleWidth      =   1065
      TabIndex        =   14
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   3
      Left            =   5280
      ScaleHeight     =   5505
      ScaleWidth      =   1065
      TabIndex        =   13
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   2
      Left            =   3840
      ScaleHeight     =   5505
      ScaleWidth      =   1065
      TabIndex        =   12
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   1
      Left            =   2400
      ScaleHeight     =   5505
      ScaleWidth      =   1065
      TabIndex        =   11
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox picCurrentCard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   600
      ScaleHeight     =   1425
      ScaleWidth      =   1065
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   2280
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblPointsGiven 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   26
      Top             =   2880
      Width           =   975
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   2280
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   2280
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2280
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   120
      Picture         =   "StackJack by Zuhab Wasim (Main).frx":030A
      Top             =   4650
      Width           =   2190
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Multiplier for Clears:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblMultiplier 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "1x"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   3570
      Width           =   375
   End
   Begin VB.Label lblCardsLeft 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "52"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cards Remaining:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0:00.0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   840
      TabIndex        =   21
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Time Elapsed:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   4050
      Width           =   735
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   5
      Left            =   8880
      TabIndex        =   19
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblPoints 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   840
      TabIndex        =   17
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Points:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   4
      Left            =   7440
      TabIndex        =   10
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Count:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Count:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Count:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Count:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Count:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Current Card:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   570
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "Start New Game"
         Enabled         =   0   'False
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuHighscore 
         Caption         =   "Show High Scores"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuRules 
         Caption         =   "Show Rules"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Zuhab Wasim
'Date: 27/05/2017
'Purpose: To demonstrate the resources and materials learned
'         throughout the year of ICS4U and to apply it affectively
'         by replicating the game StackJack (Version 2)

Option Explicit

'All Constants used in the main form
Const MAX_CARDS = 52
Const MIN_CARDS = 1
Const MAX_STACKS = 5
Const CARD_OFFSET = 18

'Constants for each value of points in the game
Const POINTS_DISCARD = -150
Const POINTS_CARDPLACED = 40
Const POINTS_SPECIALCARDPLACED = 50
Const POINTS_BUST = -700
Const POINTS_CLEAR = 500

'The deck of cards and the current card number
Dim CardCounter As Integer
Dim Cards(1 To MAX_CARDS) As Long

'Arrays pretaining to the stack columns
Dim HighValueAce(1 To MAX_STACKS) As Boolean
Dim Stacks(1 To MAX_STACKS) As Integer
Dim StackCount(1 To MAX_STACKS) As Integer
Dim StackValue(1 To MAX_STACKS) As Integer

'The score in the game
Dim Points As Integer
Dim ScoreMultiplier As Integer
Dim LastStackClear As Integer

'Timer variable
Dim TimeCount As Single

'Used to initialize card dimensions
Dim CardHeight As Long
Dim CardWidth As Long

Private Sub cmdDiscard_Click()
    
    'Starts the game
    tmrTimer.Enabled = True
    
    'If the user has discarded a card, their multiplier reverts back to 1
    lblMultiplier.Caption = "1x"
    
    'Deducts 150 points from the user
    UpdatePoints lblPoints, Points, POINTS_DISCARD, lblPointsGiven
    
    'Draws a new card from the deck
    DrawNewCard picCurrentCard, Cards(), CardCounter, Points, lblCardsLeft, MAX_CARDS
    
    'If no more cards exist, end the game
    If GameOver Then
        EndGame Points, TimeCount, cmdDiscard, picStack(), MAX_STACKS, mnuNewGame, tmrTimer
    End If
    
End Sub

Private Sub Form_Load()
    
    Dim CardInit As Long
    Dim Ret As Long
    
    'Ensures the form is seen so the cards being printed during form load can be seen
    frmMain.Show
    
    'Randomizes the seed generation, making every game unique
    Randomize
    
    'Initializes the dimensions of the card
    CardInit = cdtInit(CardWidth, CardHeight)
    
    'Initializes all variables and arrays used in the program
    Initialize Cards(), MAX_CARDS, CardCounter, Stacks(), StackCount(), StackValue(), Points, ScoreMultiplier, HighValueAce(), MAX_STACKS, TimeCount
    
    'Randomizes the card sequence
    ShuffleDeck Cards(), MAX_CARDS, MIN_CARDS
    
    'Displays a card to begin the game
    DrawNewCard picCurrentCard, Cards(), CardCounter, Points, lblCardsLeft, MAX_CARDS
    
    'Sets the initial record length of the highscores record file
    HighScoreLength = Len(HighScores(1))
    
    'If the highscores record file does not exist, create one
    If Dir$(App.Path & "/" & HIGHSCORE_FILE) = "" Then
        CreateFile
    End If
    
    'Retrieves the values from the highscore record file
    GetHighScores HighScores(), HIGHSCORE_FILE, HighScoreLength
    
    'Centers the form
    CenterForm frmMain
    
    'EndGame 9567, 463, cmdDiscard, picStack(), MAX_STACKS, mnuNewGame, tmrTimer
    
End Sub

Private Sub mnuAbout_Click()
    
    'Pauses the timer if the game is still going on
    If GamePausable Then
        tmrTimer.Enabled = False
        lblTime.Caption = "PAUSED"
    End If
    
    'Displays the about form
    frmAbout.Show vbModal
    
End Sub

Private Sub mnuExit_Click()
    
    Dim EMsg As String
    Dim EType As Integer
    Dim ETitle As String
    
    Dim EResponse As Integer
    
    'Asks the user if they would like to confirm their exit request then ends the program if so.
    EMsg = "Are you sure you want to exit?"
    EType = vbInformation + vbYesNo
    ETitle = "Exit"
    
    'If the game is going on, pause it
    If GamePausable Then
        tmrTimer.Enabled = False
        lblTime.Caption = "PAUSED"
    End If
    
    EResponse = MsgBox(EMsg, EType, ETitle)
    
    If EResponse = vbYes Then
        End
    Else
        'If the game was going on that means it was paused, resume it
        If GamePausable = True Then
            tmrTimer.Enabled = True
        End If
    End If
    
End Sub

Private Sub mnuHighscore_Click()
    
    'Pauses the timer if the game is still going on
    If GamePausable Then
        tmrTimer.Enabled = False
        lblTime.Caption = "PAUSED"
    End If
    
    'Displays the high scores form
    frmHighScores.Show vbModal
    
End Sub

Private Sub mnuNewGame_Click()
    
    mnuNewGame.Enabled = False
    tmrTimer.Enabled = False
    
    'Initializes all labels, and stack columns to start the new game
    ResetForm lblCount(), picStack(), MAX_STACKS, lblPoints, cmdDiscard, lblTime, lblMultiplier, lblPointsGiven
    
    'Initializes all variables and arrays used in the program
    Initialize Cards(), MAX_CARDS, CardCounter, Stacks(), StackCount(), StackValue(), Points, ScoreMultiplier, HighValueAce(), MAX_STACKS, TimeCount
    
    'Randomizes the card sequence
    ShuffleDeck Cards(), MAX_CARDS, MIN_CARDS
    
    'Displays a card to begin the game
    DrawNewCard picCurrentCard, Cards(), CardCounter, Points, lblCardsLeft, MAX_CARDS
    
End Sub

Private Sub mnuRules_Click()
    
    'Pauses the timer if the game is still going on
    If GamePausable Then
        tmrTimer.Enabled = False
        lblTime.Caption = "PAUSED"
    End If
    
    'Displays the rules form
    frmRules.Show vbModal
    
End Sub

Private Sub picStack_Click(Index As Integer)
    
    Dim Ret As Long
    
    tmrTimer.Enabled = True
    GamePausable = True
    
    'Prints the card onto the desired stack, offsetting down everytime so the previous
    'cards can be seen
    Ret = cdtDraw(picStack(Index).hDC, 0, CARD_OFFSET * StackCount(Index), Cards(CardCounter), C_FACES, 0)
    picStack(Index).Refresh
    
    'Adds the value of the card to the stack's counter
    StackValue(Index) = StackValue(Index) + AddValue(StackValue(Index), Cards(CardCounter), HighValueAce(Index))
    
    'If the user goes above 21 but there exists an ace in the stack valued at 11 and not 1
    'It will deduce the value of that ace back to 1, avoiding a bust
    If StackValue(Index) > 21 And HighValueAce(Index) Then
        StackValue(Index) = StackValue(Index) - 10
        HighValueAce(Index) = False
    End If
    lblCount(Index).Caption = VBA.Str$(StackValue(Index))
    
    'Initially sets the multiplier caption to only being 1
    lblMultiplier.Caption = "1x"
    
    'If the value of the stack is still above 21
    'The column is bust
    If StackValue(Index) > 21 Then
        BustStack lblCount(Index), picStack(), Index, MAX_STACKS
        'If the user has busted all of their stacks, take away all of their points
        If GameOver Then
            UpdatePoints lblPoints, Points, -Points, lblPointsGiven
        Else
            UpdatePoints lblPoints, Points, POINTS_BUST, lblPointsGiven
        End If
    'If its 21 then then it clears the stack and reinitializes back to 0
    ElseIf StackValue(Index) = 21 Then
        lblCount(Index).Caption = "0"
        StackValue(Index) = 0
        HighValueAce(Index) = False
        'Sets StackCount to -1 as it will become 0 by the end of the procedure due to the increment right after
        StackCount(Index) = -1
        picStack(Index).Cls
        UpdatePoints lblPoints, Points, StackPoints(POINTS_CLEAR, LastStackClear, CardCounter, ScoreMultiplier), lblPointsGiven
        LastStackClear = CardCounter
        'If the player gets consecutive clears, the score multiplier would go up
        'meaning the multiplier label should be updated as well
        lblMultiplier.Caption = Str$(ScoreMultiplier + 1) & "x"
    Else
        'Checks to see what type of card is placed, and awards more points to face cards and aces
        If Cards(CardCounter) >= 40 Or Cards(CardCounter) <= 3 Then
            UpdatePoints lblPoints, Points, POINTS_SPECIALCARDPLACED, lblPointsGiven
        Else
            UpdatePoints lblPoints, Points, POINTS_CARDPLACED, lblPointsGiven
        End If
    End If
    
    'Increments the number of cards in the stack
    StackCount(Index) = StackCount(Index) + 1
    
    'Draws the next card in the sequence
    DrawNewCard picCurrentCard, Cards(), CardCounter, Points, lblCardsLeft, MAX_CARDS
    
    'Checks to see if the game is over due to all stacks being bust or drawing a new card
    If GameOver Then
        EndGame Points, TimeCount, cmdDiscard, picStack(), MAX_STACKS, mnuNewGame, tmrTimer
    End If
    
End Sub

Private Sub tmrTimer_Timer()
    
    'Timer goes in miliseconds
    TimeCount = TimeCount + 0.1
    
    'Outputs the time as M:SS.S, or HH:MM:SS.S
    lblTime.Caption = ConvTime(TimeCount)
    
    'If the user reaches 59:59:59.9 for their time, they lose
    If TimeCount >= 216000 - 0.1 Then
        tmrTimer.Enabled = False
        GameOverMessage = "You lose! You have run out of time and lost any of the points you had! "
        'Take away all of their points for losing
        UpdatePoints lblPoints, Points, -Points, lblPointsGiven
        EndGame Points, TimeCount, cmdDiscard, picStack(), MAX_STACKS, mnuNewGame, tmrTimer
    End If
    
End Sub
