VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StackJack"
   ClientHeight    =   6045
   ClientLeft      =   2910
   ClientTop       =   2985
   ClientWidth     =   8820
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
   ScaleHeight     =   6045
   ScaleWidth      =   8820
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   5
      Left            =   7560
      ScaleHeight     =   5505
      ScaleWidth      =   1065
      TabIndex        =   18
      Top             =   360
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
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   1335
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   4
      Left            =   6120
      ScaleHeight     =   5505
      ScaleWidth      =   1065
      TabIndex        =   14
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   3
      Left            =   4680
      ScaleHeight     =   5505
      ScaleWidth      =   1065
      TabIndex        =   13
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   2
      Left            =   3240
      ScaleHeight     =   5505
      ScaleWidth      =   1065
      TabIndex        =   12
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   1
      Left            =   1800
      ScaleHeight     =   5505
      ScaleWidth      =   1065
      TabIndex        =   11
      Top             =   360
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
      Left            =   240
      ScaleHeight     =   1425
      ScaleWidth      =   1065
      TabIndex        =   0
      Top             =   360
      Width           =   1095
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
      Left            =   8280
      TabIndex        =   19
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblPoints 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   2520
      Width           =   615
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
      Top             =   2520
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
      Left            =   6840
      TabIndex        =   10
      Top             =   120
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
      Left            =   5400
      TabIndex        =   9
      Top             =   120
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
      Left            =   3960
      TabIndex        =   8
      Top             =   120
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
      Left            =   2520
      TabIndex        =   7
      Top             =   120
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
      Left            =   3240
      TabIndex        =   6
      Top             =   120
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
      Left            =   7560
      TabIndex        =   5
      Top             =   120
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
      Left            =   6120
      TabIndex        =   4
      Top             =   120
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
      Left            =   4680
      TabIndex        =   3
      Top             =   120
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
      Left            =   1800
      TabIndex        =   2
      Top             =   120
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
         Enabled         =   0   'False
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuHighscore 
         Caption         =   "Highscores"
         Enabled         =   0   'False
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
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuRules 
         Caption         =   "Rules"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Zuhab Wasim
'Date: 08/05/2017
'Purpose: To demonstrate the resources and materials learned
'         throughout the year of ICS4U and to apply it affectively
'         by replicating the game StackJack

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

'Used to initialize card dimensions
Dim CardHeight As Long
Dim CardWidth As Long

Private Sub cmdDiscard_Click()
    
    'Deducts 150 points from the user
    UpdatePoints lblPoints, Points, POINTS_DISCARD
    
    'Draws a new card from the deck
    DrawNewCard picCurrentCard, Cards(), CardCounter, Points
    
    'If no more cards exist, end the game
    If GameOver Then
        EndGame Points, cmdDiscard, picStack(), MAX_STACKS, mnuNewGame
    End If
    
End Sub

Private Sub Form_Load()
    
    Dim CardInit As Long
    Dim Ret As Long
    
    'Ensures the form is seen so the cards being printed during form load cna be seen
    frmMain.Show
    
    'Randomizes the seed generation, making every game unique
    Randomize
    
    'Initializes the dimensions of the card
    CardInit = cdtInit(CardWidth, CardHeight)
    
    'Initializes all variables and arrays used in the program
    Initialize Cards(), MAX_CARDS, CardCounter, Stacks(), StackCount(), StackValue(), Points, HighValueAce(), MAX_STACKS
    
    'Randomizes the card sequence
    ShuffleDeck Cards(), MAX_CARDS, MIN_CARDS
    
    'Displays a card to begin the game
    DrawNewCard picCurrentCard, Cards(), CardCounter, Points
    
End Sub

Private Sub mnuAbout_Click()
    
    'Displays the about form
    frmAbout.Show vbModal
    
End Sub

Private Sub mnuExit_Click()
    
    Dim EMsg As String
    Dim EType As Integer
    Dim ETitle As String
    
    Dim EResponse As Integer
    
    'Asks the user if they would like to confirm their exit request then ends the program.
    EMsg = "Are you sure you want to exit?"
    EType = vbInformation + vbYesNo
    ETitle = "Exit"
    
    EResponse = MsgBox(EMsg, EType, ETitle)
    
    If EResponse = vbYes Then
        End
    End If
    
End Sub

Private Sub mnuNewGame_Click()
    
    mnuNewGame.Enabled = False
    
    'Initializes all labels, and stack columns to start the new game
    ResetForm lblCount(), picStack(), MAX_STACKS, lblPoints, cmdDiscard
    
    'Initializes all variables and arrays used in the program
    Initialize Cards(), MAX_CARDS, CardCounter, Stacks(), StackCount(), StackValue(), Points, HighValueAce(), MAX_STACKS
    
    'Randomizes the card sequence
    ShuffleDeck Cards(), MAX_CARDS, MIN_CARDS
    
    'Displays a card to begin the game
    DrawNewCard picCurrentCard, Cards(), CardCounter, Points
    
End Sub

Private Sub mnuRules_Click()
    
    'Displays the rules form
    frmRules.Show vbModal
    
End Sub

Private Sub picStack_Click(Index As Integer)
    
    Dim Ret As Long
    
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
    lblCount(Index).Caption = Str$(StackValue(Index))
    
    'If the value of the stack is still above 21
    'The column is bust
    If StackValue(Index) > 21 Then
        BustStack lblCount(Index), picStack(), Index, MAX_STACKS, Points
        UpdatePoints lblPoints, Points, POINTS_BUST
    'If its 21 then then it clears the stack and reinitializes back to 0
    ElseIf StackValue(Index) = 21 Then
        lblCount(Index).Caption = "0"
        StackValue(Index) = 0
        HighValueAce(Index) = False
        'Sets StackCount to -1 as it will become 0 by the end of the procedure due to the increment right after
        StackCount(Index) = -1
        picStack(Index).Cls
        UpdatePoints lblPoints, Points, POINTS_CLEAR
    Else
        'Checks to see what type of card is placed, and awards more points to face cards and aces
        If Cards(CardCounter) >= 39 Or Cards(CardCounter) <= 3 Then
            UpdatePoints lblPoints, Points, POINTS_SPECIALCARDPLACED
        Else
            UpdatePoints lblPoints, Points, POINTS_CARDPLACED
        End If
    End If
    
    'Increments the number of cards in the stack
    StackCount(Index) = StackCount(Index) + 1
    
    'Draws the next card in the sequence
    DrawNewCard picCurrentCard, Cards(), CardCounter, Points
    
    'Checks to see if the game is over due to all stacks being bust or drawing a new card
    If GameOver Then
        EndGame Points, cmdDiscard, picStack(), MAX_STACKS, mnuNewGame
    End If
    
End Sub
