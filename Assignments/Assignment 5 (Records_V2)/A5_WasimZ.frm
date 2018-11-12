VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Random Access"
   ClientHeight    =   7470
   ClientLeft      =   2880
   ClientTop       =   2760
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "A5_WasimZ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   6945
   Begin VB.ListBox lstData 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   6810
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6735
   End
   Begin MSComDlg.CommonDialog cdlDialog 
      Left            =   4200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picData 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   360
      Width           =   6735
   End
   Begin VB.Label lblRecords 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   15
      Width           =   1815
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
         Shortcut        =   {DEL}
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
'Date: 19/01/2017
'Purpose: To practice with the manipulation of data from record files and converting from text files to record files.
'         To also practice using listboxes in Visual Basic 6.0
Option Explicit

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

    'Prints the header and dash-line seperator
    picData.Print "  #."; Tab(6); _
                 "LastName"; Tab(27); _
                 "FirstName"; Tab(44); _
                 "HF"; Tab(54); _
                 "Mark(%)"
                 
End Sub

Private Sub lstData_Click()
    
    'Enables the modify and delete menu buttons if an item is selected
    mnuModRec.Enabled = True
    mnuDelRec.Enabled = True
    
End Sub

Private Sub lstData_DblClick()
    
    'Opens the changer form as a modifier from double click
    If lstData.ListIndex <> -1 Then
        FormType = True
        frmChanger.Show vbModal
    End If
    
End Sub

'Loads the about form
Private Sub mnuAbout_Click()
    
    'Makes it so that the about form is the focus and nothing else can be clicked
    frmAbout.Show vbModal
    
End Sub

Private Sub mnuAddRec_Click()
    
    'Checks to see if items exceed the max number of items
    If NumStudents < MAX Then
        'Sets the form to load in add format and displays it
        FormType = False
        frmChanger.Show vbModal
    Else
        MsgBox "Error: Maximum number of list items entered!", vbExclamation + vbOKOnly, "Error"
    End If
    
End Sub

Private Sub mnuDelRec_Click()
    
    If lstData.ListIndex <> -1 Then
        'Confirms the deletion of an item
        If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion, "Confirm Deletion") = vbYes Then
            
            'Calls the delete procedure and disables
            DeleteRecord lstData, Student(), NumStudents
            mnuDelRec.Enabled = False
            mnuModRec.Enabled = False
            
            'Correct grammar
            If NumStudents = 1 Then
                frmMain.lblRecords.Caption = VBA.Str$(NumStudents) & " record."
            Else
                frmMain.lblRecords.Caption = VBA.Str$(NumStudents) & " records."
            End If
        End If
    End If
    
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

Private Sub mnuModRec_Click()
    
    'Calls the changer form in modify format
    If lstData.ListIndex <> -1 Then
        FormType = True
        
        frmChanger.Show vbModal
    End If
    
End Sub

'Asks for the file the user wants to open and reads the data into the student record array and displays it
Private Sub mnuOpen_Click()
        
    'Gets file from the dialog box
    FileName = GetFile(cdlDialog)
    
    'Checks to see if the user has selected a file
    If FileName <> "" Then
        'Reads the file
        ReadFile Student(), FileName, RecLength, NumStudents
        'Checks to see if the file has any elements to display
        If NumStudents > 0 Then
            DisplayAll lstData, lblRecords, Student(), NumStudents
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

