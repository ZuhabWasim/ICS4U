VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Day of the Week"
   ClientHeight    =   2415
   ClientLeft      =   4185
   ClientTop       =   3735
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "A4_WasimZ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   5715
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdFindDay 
      Caption         =   "Get &Day of the Week"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   3975
   End
   Begin VB.ComboBox cboYear 
      Height          =   405
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.ComboBox cboDay 
      Height          =   405
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox cboMonth 
      Height          =   405
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblWeekDay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   "Year:"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Day:"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Month:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Zuhab Wasim 12E
'Date: 17/01/2017
'Purpose: To apply the knowledge learned of the Visual Basic list/combo box controls
'         by creating a program that returns the weekday based on month, day and year using Zeller's algorithm
Option Explicit

'Constants that represent values that can be changed in the future
Const MAX_MONTHS = 12
Const MAX_DAYS = 31
Const MAX_YEARS = 200
Const START_YEAR = 1900

'Variables to hold the date's month, day and year
Dim CMonth As Integer
Dim CDay As Integer
Dim CYear As Integer

Private Sub cmdExit_Click()
    
    'Validates the user's click on ending the program
    If MsgBox("Are you sure you want to exit?", vbYesNo + vbInformation, "Exit") = vbYes Then
        End
    End If
    
End Sub

Private Sub Form_Load()
    
    'Calls the Initialize procedure to set up the combo boxes
    Initialize MAX_MONTHS, MAX_DAYS, MAX_YEARS, START_YEAR

End Sub

Public Sub cmdFindDay_Click()
    
    CMonth = Val(cboMonth.ListIndex + 1)
    CDay = cboDay.ListIndex + 1
    CYear = cboYear.ListIndex + START_YEAR
    
    'Checks if date is valid
    If ValidDate(CDay, CMonth, CYear) Then
        lblWeekDay.Caption = GetWeekDay(CMonth, CDay, CYear)
    Else
        MsgBox "Sorry, that is not a valid date!", vbOKOnly + vbExclamation, "Error: Invalid Date"
        'Resets the day outputted if date is not valid
        lblWeekDay.Caption = ""
    End If
            
End Sub
