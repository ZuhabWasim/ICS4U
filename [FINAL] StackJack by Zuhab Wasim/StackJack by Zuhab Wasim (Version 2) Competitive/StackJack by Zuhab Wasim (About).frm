VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0FFC0&
   Caption         =   "About"
   ClientHeight    =   5175
   ClientLeft      =   3645
   ClientTop       =   2385
   ClientWidth     =   6165
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StackJack by Zuhab Wasim (About).frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6165
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Return"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Version: 2.0.0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "© Z. Wasim 2017, All Rights Reserved."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
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
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Aim for the highest score and good luck!"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   5895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   $"StackJack by Zuhab Wasim (About).frx":030A
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   $"StackJack by Zuhab Wasim (About).frx":03C6
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   5655
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5520
      Picture         =   "StackJack by Zuhab Wasim (About).frx":0490
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "StackJack by Zuhab Wasim (About).frx":079A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "About the game"
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
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReturn_Click()
    
    'Unloads the form and returns back to the game
    Unload frmAbout

End Sub

Private Sub Form_Load()
    
    'Centers the form
    CenterForm frmAbout
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Enabled the timer if the game is still going on
    If GamePausable Then
        frmMain.tmrTimer.Enabled = True
    End If
    
End Sub
