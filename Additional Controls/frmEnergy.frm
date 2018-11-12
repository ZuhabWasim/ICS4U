VERSION 5.00
Begin VB.Form frmEnergy 
   Caption         =   "Energy Calculator"
   ClientHeight    =   3765
   ClientLeft      =   9495
   ClientTop       =   5205
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   4275
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Estimated Cost"
      Height          =   975
      Left            =   1920
      TabIndex        =   10
      Top             =   1080
      Width           =   2295
      Begin VB.Label lblCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$0.00"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Kilowatt Hours"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   4095
      Begin VB.HScrollBar hsbHours 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblHours 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Season"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton optWinter 
         Caption         =   "Winter"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optFall 
         Caption         =   "Fall"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optSummer 
         Caption         =   "Summer"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optSpring 
         Caption         =   "Spring"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmEnergy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const KWHRATE = 0.07
Dim SeasonRate As Single

Private Sub cmdClear_Click()
    
    optSpring.Value = True
    hsbHours.Value = 0
    hsbHours.SetFocus
    
End Sub

Private Sub cmdExit_Click()

    End
    
End Sub

Private Sub NewCost()

    Dim KwHours As Integer
    
    KwHours = hsbHours.Value
    lblCost.Caption = Format$(KwHours * (KWHRATE + SeasonRate), "$##,##0.00")
    
End Sub

Private Sub cmdReturn_Click()
    
    Unload frmEnergy
    frmHub.Show
    
End Sub

Private Sub Form_Load()

    optSpring.Value = True
    
End Sub

Private Sub optWinter_Click()
    
    SeasonRate = 0.02
    NewCost
    
End Sub

Private Sub optSummer_Click()
    
    SeasonRate = 0.02
    NewCost
    
End Sub

Private Sub optSpring_Click()
    
    SeasonRate = 0.01
    NewCost
    
End Sub

Private Sub optFall_Click()
    
    SeasonRate = 0.01
    NewCost
    
End Sub

Private Sub lblHours_Change()

    NewCost
    
End Sub

Private Sub hsbHours_Scroll()

    lblHours.Caption = Str$(hsbHours.Value)
    
End Sub

Private Sub hsbHours_Change()

    lblHours.Caption = Str$(hsbHours.Value)
    
End Sub
