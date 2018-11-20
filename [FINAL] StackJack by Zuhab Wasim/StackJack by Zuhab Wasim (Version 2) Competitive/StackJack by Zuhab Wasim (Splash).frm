VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   2865
   ClientLeft      =   3825
   ClientTop       =   3960
   ClientWidth     =   7755
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StackJack by Zuhab Wasim (Splash).frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrSplash 
      Interval        =   3000
      Left            =   7080
      Top             =   1560
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Any replication or unauthorized use of the following game created will result in legal action."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Warning:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "© Z. Wasim 2017, All Rights Reserved."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   2580
      Left            =   120
      Picture         =   "StackJack by Zuhab Wasim (Splash).frx":030A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2700
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "by: Zuhab Wasim"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "StackJack"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    'Centers the form
    CenterForm frmSplash
    
End Sub

Private Sub tmrSplash_Timer()
    
    'Unloads this form after 3 seconds of delay
    Unload frmSplash
    
    'Loads the game form
    Load frmMain
    
End Sub
