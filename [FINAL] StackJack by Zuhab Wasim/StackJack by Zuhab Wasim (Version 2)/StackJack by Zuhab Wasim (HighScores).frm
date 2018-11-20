VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHighScores 
   BackColor       =   &H00FFC0FF&
   Caption         =   "High Scores"
   ClientHeight    =   8250
   ClientLeft      =   4440
   ClientTop       =   1800
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StackJack by Zuhab Wasim (HighScores).frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   7440
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grdHighScores 
      Height          =   5880
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   10372
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6000
      Picture         =   "StackJack by Zuhab Wasim (HighScores).frx":030A
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "StackJack by Zuhab Wasim (HighScores).frx":0614
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "StackJack -"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "High Scores"
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
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "These are the highest scores acheived by the past players who have played this game!"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   5655
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constants used for adjusting the width and height of the grid control
Const COL_ADJUST = 97.5
Const ROW_ADJUST = -10

Private Sub cmdReturn_Click()
    
    'Unloads the form
    Unload frmHighScores
    
End Sub

Private Sub Form_Load()
    
    Dim Rw As Integer
    Dim Cl As Integer
    Dim RHeight As Integer
    
    CenterForm frmHighScores
    
    With grdHighScores
    
        'Sets initial properties
        .FixedRows = 1
        .FixedCols = 0
        .Rows = HIGHSCORE_MAX + 1
        .Cols = 4
        
        'Sets the width and height properties
        'Each column is a fraction of the width
        'Slight subtraction to avoid grid control scroll bar
        .ColWidth(0) = .Width * (8 / 128)
        .ColWidth(1) = .Width * (60 / 128) - 100
        .ColWidth(2) = .Width * (30 / 128) - 6
        .ColWidth(3) = .Width * (30 / 128) - 6
        RHeight = .Height / 11
        For Rw = 0 To .Rows - 1
            .RowHeight(Rw) = RHeight
        Next Rw
        
        'Sets the column headers
        .Row = 0
        .Col = 0
        .Text = " #."
        .Col = 1
        .Text = "          Name"
        .Col = 2
        .Text = "    Score"
        .Col = 3
        .Text = "    Time"
        
        'Assigns the fields of each highscore into the cells
        For Rw = 1 To .Rows - 1
            .Row = Rw
            .Col = 0
            .Text = VBA.Format$(Rw, "#0.") & ""
            .Col = 1
            .Text = HighScores(Rw).HName
            .Col = 2
            .Text = VBA.Format$(HighScores(Rw).HScore, "#,##0")
            .Col = 3
            .Text = ConvTime(HighScores(Rw).HTime)
        Next Rw
        
        'Aligns the ends of the grid control's height and width with the grid itself
        .Width = .Width * (8 / 128) + .Width * (60 / 128) + .Width * (30 / 128) + .Width * (30 / 128) + ROW_ADJUST
        .Height = .RowHeight(1) * .Rows + COL_ADJUST
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Enabled the timer if the game is still going on
    If GamePausable Then
        frmMain.tmrTimer.Enabled = True
    End If
    
End Sub
