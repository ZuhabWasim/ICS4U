VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFC0&
   Caption         =   "StackJack"
   ClientHeight    =   6045
   ClientLeft      =   1890
   ClientTop       =   1905
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   1080
      TabIndex        =   20
      Top             =   4920
      Width           =   375
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
      TabIndex        =   17
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
      Left            =   7560
      ScaleHeight     =   5505
      ScaleWidth      =   1065
      TabIndex        =   16
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
      Left            =   6120
      ScaleHeight     =   5505
      ScaleWidth      =   1065
      TabIndex        =   15
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
      Left            =   4680
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
      Index           =   1
      Left            =   3240
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
      Index           =   0
      Left            =   1800
      ScaleHeight     =   5505
      ScaleWidth      =   1065
      TabIndex        =   12
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
   Begin VB.Label lblPoints 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   19
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
      TabIndex        =   18
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
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
      Index           =   4
      Left            =   8280
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
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
      Index           =   3
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
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

Const MAX_CARDS = 52
Const MIN_CARDS = 1

Dim CardCounter As Integer
Dim Cards(1 To MAX_CARDS) As Long

Dim CardHeight As Long
Dim CardWidth As Long

Private Sub Command1_Click()
    
    Dim Ret As Long
    
    CardCounter = CardCounter + 1
    Ret = cdtDraw(picCurrentCard.hDC, 0, 0, Cards(CardCounter), C_FACES, 0)
    
    picCurrentCard.Refresh
    
End Sub

Private Sub Form_Load()
    
    Dim CardInit As Long
    
    frmMain.Show
    
    Randomize
     
    CardInit = cdtInit(CardWidth, CardHeight)
    
    Initialize Cards(), MAX_CARDS, CardCounter
    ShuffleDeck Cards(), MAX_CARDS, MIN_CARDS
    
    picCurrentCard.Refresh
    
End Sub

