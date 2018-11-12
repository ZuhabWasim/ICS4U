VERSION 5.00
Begin VB.Form frmFileInfo 
   Caption         =   "File Information"
   ClientHeight    =   6435
   ClientLeft      =   9525
   ClientTop       =   5265
   ClientWidth     =   8520
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
   ScaleHeight     =   6435
   ScaleWidth      =   8520
   Begin VB.TextBox txtFilename 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   1695
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   3495
   End
   Begin VB.DirListBox dirDirectory 
      Height          =   4680
      Left            =   3720
      TabIndex        =   7
      Top             =   840
      Width           =   4695
   End
   Begin VB.FileListBox filFile 
      Height          =   1890
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   3495
   End
   Begin VB.ComboBox cboFileType 
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3720
      Width           =   3495
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblDir 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label4 
      Caption         =   "Drive:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "File Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Directories:"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "File Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReturn_Click()
    
    Unload frmFileInfo
    frmHub.Show
    
End Sub

Private Sub Form_Load()
    
    With cboFileType
        .AddItem "All Files (*.*)"
        .AddItem "Executable Files (*.exe)"
        .AddItem "Text Files (*.txt)"
        
        .ListIndex = 0
    End With
    
    lblDir.Caption = dirDirectory.Path
    
End Sub

Private Sub cmdOK_Click()

    If txtFilename.Text = "" Then
        MsgBox "Select a filename first!", vbExclamation, "Select File"
    Else
        GetFileInfo
    End If
    
End Sub

Private Sub cmdCancel_Click()
    
    End
    
End Sub

Private Sub cboFileType_Click()
    
    Select Case cboFileType.ListIndex
        Case 0
            filFile.Pattern = "*.*"
        Case 1
            filFile.Pattern = "*.exe"
        Case 2
            filFile.Pattern = "*.txt*"
    End Select
    
End Sub

Private Sub dirDirectory_Change()

    filFile.Path = dirDirectory.Path
    lblDir.Caption = dirDirectory.Path
    txtFilename.Text = ""
    
End Sub

Private Sub drvDrive_Change()

    On Error GoTo DriveError
    dirDirectory.Path = drvDrive.Drive
    Exit Sub
    
DriveError:
    MsgBox "A drive error occurred!", vbExclamation, "Drive Error"
    drvDrive.Drive = dirDirectory.Path
    Exit Sub
End Sub

Private Sub filFile_Click()
    
    txtFilename.Text = filFile.FileName
    
End Sub

Private Sub filFile_dblClick()

    txtFilename.Text = filFile.FileName
    cmdOK_Click
    
End Sub

Private Sub GetFileInfo()

    Dim Path As String
    Dim FName As String
    Dim FileInfo As String
    
    Path = filFile.Path
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    FName = Path & filFile.FileName
    
    On Error GoTo FileError
    FileInfo = "Last Modified On: " & FileDateTime(FName) & vbCrLf & _
               "File Size: " & Format$(FileLen(FName), "#,###,###") & " bytes"
    MsgBox FileInfo, vbInformation, UCase$(filFile.FileName)
    Exit Sub
    
FileError:
    MsgBox "An error occurred while retrieving " & FName & "'s information!", _
            vbExclamation, "File Information Error"
End Sub
