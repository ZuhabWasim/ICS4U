VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Random Access"
   ClientHeight    =   7695
   ClientLeft      =   4170
   ClientTop       =   2490
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "A3_WasimZ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   6630
   Begin MSComDlg.CommonDialog cdlDialog 
      Left            =   240
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picData 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   120
      ScaleHeight     =   7035
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "Total Students:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label lblRecords 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   7200
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuAddRec 
         Caption         =   "&Add Record..."
      End
      Begin VB.Menu mnuDelRec 
         Caption         =   "&Delete Record"
      End
      Begin VB.Menu mnuModRec 
         Caption         =   "&Modify Record..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "A&bout"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Zuhab Wasim
'Date: 10/11/2016
'Purpose: To practice with the manipulation of data from record files and converting from text files to record files.

Option Explicit

Const MAX = 100

Dim Student(1 To MAX) As StudentRec
Dim NumStudents As Integer
Dim RecLength As Integer

Dim FileName As String

'The code executed when the form is loaded (when the program starts as this is the main form loaded)
Private Sub Form_Load()
    
    'Initializes the number of students and the length of the record
    NumStudents = 0
    RecLength = Len(Student(1))
    
    'Disables the features not used in this version
    mnuDelRec.Enabled = False
    mnuModRec.Enabled = False
                          
End Sub

'Loads the about form
Private Sub mnuAbout_Click()
    
    'Makes it so that the about form is the focus and nothing else can be clicked
    frmAbout.Show vbModal
    
End Sub

Private Sub mnuAddRec_Click()
    
    'Conveys to the user that the feature clicked is currently unavailable
    MsgBox "This feature is currently unavailable and will be implemented in the next version of this program.", vbOKOnly + vbInformation, "In Progress"
    
End Sub

'Exits the program after confirming the user's click
Private Sub mnuExit_Click()
    
    Dim EMsg As String
    Dim EType As Integer
    Dim ETitle As String
    
    Dim EResponse As Integer
    
    EMsg = "Are you sure you want to exit"
    EType = vbInformation + vbYesNo
    ETitle = "Exit"
    
    EResponse = MsgBox(EMsg, EType, ETitle)
    
    If EResponse = vbYes Then
        End
    End If
    
End Sub

'Asks for the file the user wants to open and reads the data into the student record array
'and displays it onto the main form's picturebox
Private Sub mnuOpen_Click()
    
    FileName = GetFile(cdlDialog)
    
    'Checks to see if the user has selected a file
    If FileName <> "" Then
        'Reads the file
        ReadFile Student(), FileName, RecLength, NumStudents
        'Checks to see if the file has any elements to display
        If NumStudents > 0 Then
            DisplayAll picData, lblRecords, Student(), NumStudents
        End If
    End If
    
End Sub

'Asks for the filename the user wants to save as and saves a file with that name
Private Sub mnuSave_Click()
    
    'Gets the filename used for saving
    FileName = GetSaveFile(cdlDialog)
    
    'Saves the file
    SaveFile Student(), FileName, NumStudents, RecLength
    
End Sub
