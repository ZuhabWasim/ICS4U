VERSION 5.00
Begin VB.Form frmRules 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Rules"
   ClientHeight    =   7230
   ClientLeft      =   4065
   ClientTop       =   2430
   ClientWidth     =   8160
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StackJack by Zuhab Wasim (Rules).frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8160
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   6600
      Width           =   1455
   End
   Begin VB.ListBox lstRules 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   7695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "These are the rules that define how the game is played as well as the specific scoring for each action."
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   5655
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   7440
      Picture         =   "StackJack by Zuhab Wasim (Rules).frx":030A
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "StackJack by Zuhab Wasim (Rules).frx":0614
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Rules"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "StackJack -"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReturn_Click()
    
    'Unloads the form and returns back to the game
    Unload frmRules
    
End Sub

Private Sub Form_Load()
    
    'Prints out every rule on the list
    With lstRules
        .ListIndex = -1
        .AddItem " Instructions--------------------------------------------------------------------"
        .AddItem "  -To place a card, simply click on the desired stack."
        .AddItem "  -To discard a card, click on the Discard button under the current card."
        .AddItem "  -The current card image displays the current card to be placed."
        .AddItem "  -A stack with a red value on top of it indicates that the stack is bust."
        .AddItem "  -A stack with a purple value indicates you can still place cards on that stack."
        .AddItem "  -The points underneath the current card indicates the current score obtained."
        .AddItem "  -Once a highscore is acheived, enter your name to be put on the board."
        .AddItem "  -To start a new game, go to File and then click on New Game."
        .AddItem "  -To end the game, go to File and then click on Exit."
        .AddItem ""
        .AddItem " Rules"
        .AddItem "  1. You must not go over the value of 21 on any stack."
        .AddItem "  2. If you are to go over the value 21, that stack will go bust"
        .AddItem "      and you cannot use that stack for the rest of the game."
        .AddItem "  3. The game will be won if all cards are used and the score is not zero."
        .AddItem "  4. The game will be lost if the final score is zero after using all cards or"
        .AddItem "      all five stacks go bust."
        .AddItem "  5. The player cannot start a new game until the current game is completed."
        .AddItem "  6. If a highscore is acheived, the player will be put on the highscore board."
        .AddItem "  7. The player cannot go below zero for their score."
        .AddItem ""
        .AddItem " Scoring"
        .AddItem "  -Place a card successfully .......................   40 points"
        .AddItem "  -Bonus points for placing a face card or an ace ..   10 points"
        .AddItem "  -Successfully clearing a stack (value of 21) .....  500 points"
        .AddItem "  -Have a stack go bust (value over 21) ............ -700 points"
        .AddItem "  -Discarding a card ............................... -150 points"
        .AddItem ""
        
    End With
    
End Sub

Private Sub lstRules_Click()
    
    lstRules.ListIndex = -1

End Sub
