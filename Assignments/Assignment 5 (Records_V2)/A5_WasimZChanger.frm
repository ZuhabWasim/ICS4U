VERSION 5.00
Begin VB.Form frmChanger 
   Caption         =   "Detailed Information"
   ClientHeight    =   2760
   ClientLeft      =   4560
   ClientTop       =   4110
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "A5_WasimZChanger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4005
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Modify"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Return"
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtMark 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtHF 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtFirstName 
      Height          =   375
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   5
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtLastName 
      Height          =   375
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Mark:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "HF:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "First Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    
    Dim NewStudent As StudentRec
    
    'Retrieves values from text box
    With NewStudent
        .LastName = VBA.Trim$(txtLastName.Text)
        .FirstName = VBA.Trim$(txtFirstName.Text)
        .HomeForm = VBA.Trim$(txtHF.Text)
        .Mark = Val(txtMark.Text)
    End With
    
    'Checks to see if input is valid
    If Validate(NewStudent, txtMark) Then
        If NumStudents < MAX Then
            'Adds all the new student info and appends it to the array and list box
            NumStudents = NumStudents + 1
            With Student(NumStudents)
                .LastName = NewStudent.LastName
                .FirstName = NewStudent.FirstName
                .HomeForm = NewStudent.HomeForm
                .Mark = NewStudent.Mark
                
                frmMain.lstData.AddItem " " & _
                                       VBA.Format$(NumStudents, "@@") & ". " & _
                                       .LastName & " " & _
                                       .FirstName & " " & _
                                       VBA.Format$(VBA.Trim$(.HomeForm), "@@@") & "          " & _
                                       VBA.Format$(.Mark, "@@@")
            End With
            
            'Correct grammar
            If NumStudents = 1 Then
                frmMain.lblRecords.Caption = VBA.Str$(NumStudents) & " record."
            Else
                frmMain.lblRecords.Caption = VBA.Str$(NumStudents) & " records."
            End If
            
            'Clears the info entered for new fields
            txtLastName.Text = ""
            txtFirstName.Text = ""
            txtHF.Text = ""
            txtMark.Text = ""
        Else
            MsgBox "Error: Maximum number of list items entered!", vbExclamation + vbOKOnly, "Error"
        End If
    End If
    
    'Sets the textfocus to the first text box
    txtLastName.SetFocus
    
End Sub

Private Sub cmdModify_Click()

    Dim NewStudent As StudentRec
    
    'Retrieves the values from the text boxes
    With NewStudent
        .LastName = VBA.Trim$(txtLastName.Text)
        .FirstName = VBA.Trim$(txtFirstName.Text)
        .HomeForm = VBA.Trim$(txtHF.Text)
        .Mark = Val(txtMark.Text)
    End With
    
    'Checks to see if input is valid
    If Validate(NewStudent, txtMark) Then
        
        'Replaces the new student info with the selected item in the array and list box
        With Student(NumStudents)
            .LastName = NewStudent.LastName
            .FirstName = NewStudent.FirstName
            .HomeForm = NewStudent.HomeForm
            .Mark = NewStudent.Mark

            frmMain.lstData.List(frmMain.lstData.ListIndex) = " " & _
                                 VBA.Format$(frmMain.lstData.ListIndex + 1, "@@") & ". " & _
                                 .LastName & " " & _
                                 .FirstName & " " & _
                                 VBA.Format$(VBA.Trim$(.HomeForm), "@@@") & "          " & _
                                 VBA.Format$(.Mark, "@@@")
        End With

        'Correct grammar
        If NumStudents = 1 Then
            frmMain.lblRecords.Caption = VBA.Str$(NumStudents) & " record."
        Else
            frmMain.lblRecords.Caption = VBA.Str$(NumStudents) & " records."
        End If
        
        'Unselects the item that was modified
        frmMain.lstData.ListIndex = -1
        
        'Exits the form to allow the user to see the new changes
        Unload frmChanger
    End If
    
End Sub

Private Sub cmdReturn_Click()
    
    'Clears the selected item
    frmMain.lstData.ListIndex = -1
    
    'Disables menu and delete records
    frmMain.mnuDelRec.Enabled = False
    frmMain.mnuModRec.Enabled = False
    
    Unload frmChanger
    
End Sub

Private Sub Form_Load()
    
    'For the add form
    If FormType = False Then
        cmdAdd.Visible = True
        cmdModify.Visible = False
    'For the modify form
    Else
        cmdAdd.Visible = False
        cmdModify.Visible = True
        
        'Inputs the selected student's info into the field text boxes
        With Student(frmMain.lstData.ListIndex + 1)
            txtLastName.Text = VBA.Trim$(.LastName)
            txtFirstName.Text = VBA.Trim$(.FirstName)
            txtHF.Text = VBA.Trim$(.HomeForm)
            txtMark.Text = VBA.Trim$(.Mark)
        End With
    End If
    
End Sub

Private Sub txtHF_KeyPress(KeyAscii As Integer)
    
    'Reassigns a letter's ANSI value to the capitilized letter ANSI
    KeyAscii = VBA.Asc(VBA.UCase$(VBA.Chr$(KeyAscii)))
    
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z"))) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtMark_KeyPress(KeyAscii As Integer)
    
    'Only digits and backspace are allowed
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub
