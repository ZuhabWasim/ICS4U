VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Credit Card Test Analyzer"
   ClientHeight    =   6390
   ClientLeft      =   5025
   ClientTop       =   2280
   ClientWidth     =   6870
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
   ScaleHeight     =   6390
   ScaleWidth      =   6870
   Begin VB.Frame Frame1 
      Caption         =   "Results"
      Height          =   2775
      Left            =   4320
      TabIndex        =   6
      Top             =   1320
      Width           =   2415
      Begin VB.Label Label3 
         Caption         =   "Invalid Credit Cards:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblInvalid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblVisa 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblMasterCard 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblAmex 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   120
         Picture         =   "READER_WasimZ.frx":0000
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   120
         Picture         =   "READER_WasimZ.frx":49B7
         Stretch         =   -1  'True
         Top             =   960
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   120
         Picture         =   "READER_WasimZ.frx":17975
         Stretch         =   -1  'True
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.PictureBox picData 
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5355
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   840
      Width           =   4095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "&Read File && Analyze"
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label lblFileName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   6615
   End
   Begin VB.Label lblTotalCards 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Total Cards:"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Current File:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Author: Zuhab Wasim
'Date: 23/09/16
'Purpose: To demonstrate the knowledge and understanding of credit card validation
'         using Luhn's algoritm, and for students to review their skills in Visual Basic
'         for the upcoming year of ICS4U1.

Private Sub cmdRead_Click()
    
    'Constant declared for the maximum amount of input from the needed text file
    Const MAX = 20
    
    'CreditCards array declared for storing credit card numbers
    Dim CreditCards(1 To MAX) As String
    
    'Credit card counter variables declared
    Dim AmexCount As Integer
    Dim MasterCount As Integer
    Dim VisaCount As Integer
    Dim NValid As Integer
    Dim TotalCards As Integer
    
    'FileName variable declared for the file name
    Dim FileName As String
    
    'File name path stored into variable
    FileName = App.Path & "\CC_TEST.txt"
    
    'Initialization of the counter variables
    TotalCards = 0
    AmexCount = 0
    MasterCount = 0
    VisaCount = 0
    NValid = 0
    
    'Clears the board
    picData.Cls
    
    'Reads the file and stores the credit cards in an array, and counts the total number of cards
    ReadFile FileName, CreditCards, TotalCards
    
    'Determines which card number belongs to which credit card type as well as its validity
    ValidateCards CreditCards, TotalCards, AmexCount, MasterCount, VisaCount, NValid
    
    'Displays the credit cards from the text file
    Display CreditCards(), TotalCards
    
    'Displays the credit count card variable on the form and file name
    lblAmex.Caption = AmexCount
    lblMasterCard.Caption = MasterCount
    lblVisa.Caption = VisaCount
    lblInvalid.Caption = NValid
    lblTotalCards.Caption = TotalCards
    
    'Checks to see if the file name is greather than 70 characters long, cuts it off if so
    If Len(FileName) > 70 Then
        lblFileName.Caption = Left$(FileName, 70) & "..."
    Else
        lblFileName.Caption = FileName
    End If
    
End Sub

Private Sub cmdExit_Click()
    
    'Asks if the user wants to exit, and does if so
    Dim EMsg As String
    Dim ETitle As String
    
    Dim EType As Integer
    Dim EResponse As Integer
    
    EMsg = "Are you sure you want to exit?"
    EType = vbYesNo + vbInformation
    ETitle = "Exit Program"
    
    EResponse = MsgBox(EMsg, EType, ETitle)
    
    If EResponse = vbYes Then
        End
    End If
    
End Sub
