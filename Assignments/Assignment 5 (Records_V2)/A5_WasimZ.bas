Attribute VB_Name = "Module1"
Option Explicit

'Declares the student record
Type StudentRec
    LastName As String * 20
    FirstName As String * 15
    HomeForm As String * 3
    Mark As Integer
End Type

'Global constants and values used in the program
Global Const MAX = 100

'True: Modify format, False: Add format of the changer form
Global FormType As Boolean

Global NumStudents As Integer
Global Student(1 To MAX) As StudentRec

'Reads the file given from the user
Public Sub ReadFile(Record() As StudentRec, ByVal FileName As String, ByVal RecordLen As Integer, NumStudents As Integer)
    
    Dim K As Integer
    
    K = 0
    
    'The code used to open the file as a record file (.REC)
    If VBA.UCase$(VBA.Right$(FileName, 4)) = ".REC" Then
        
        Open FileName For Random As #1 Len = RecordLen
    
        Do While Not EOF(1)
            K = K + 1
            'Retrieves a byte length of RecordLen into each element of the student record array
            Get #1, K, Record(K)
        Loop
        
        Close #1
        
        NumStudents = K - 1
    
    'The code used to open the file as a text file (.TXT)
    ElseIf VBA.UCase$(VBA.Right$(FileName, 4)) = ".TXT" Then
    
        Open FileName For Input As #1
    
        Do While Not EOF(1)
            K = K + 1
            'Retrieves 4 values seperated by commas, into the student record array
            With Record(K)
                Input #1, .LastName
                Input #1, .FirstName
                Input #1, .HomeForm
                Input #1, .Mark
            End With
        Loop
        
        Close #1
        
        'Sets the actual number of items into the NumStudents variable
        NumStudents = K
        
    End If
    
        
End Sub

'Displays the contents of the student record into the main form's picture box and label
Public Sub DisplayAll(ByVal lstBox As Control, _
                      ByRef lblTotal As Control, _
                      ByRef Student() As StudentRec, _
                      ByVal NumStudents As Integer)

    Dim X As Integer
    
    'Clears the list box of previous content
    lstBox.Clear
    
    'Prints each field of the student record of the array
    For X = 1 To NumStudents
        lstBox.AddItem VBA.Format$(X, "@@@") & ". " & Student(X).LastName & " " & Student(X).FirstName & " " & _
                       VBA.Format$(VBA.Trim$(Student(X).HomeForm), "@@@") & "          " & _
                       VBA.Format$(Student(X).Mark, "@@@")
    Next X
    
    'Displays the actual number of records
    lblTotal.Caption = VBA.Str$(NumStudents) & " records."
    
End Sub

'Saves the elements/fields of the student record into a record fil
Public Sub SaveFile(ByRef Student() As StudentRec, ByVal FileName As String, ByVal NumStudents As Integer, ByVal RecordLen As Integer)

    Dim X As Integer
    
    'Deletes the file if it already exists, and skips the error if does not exist
    On Error GoTo ErrorHandler
    Kill FileName
    
    'Opens the file for outputing to a record file
    Open FileName For Random As #1 Len = RecordLen
    
    'Saves each element of the student record array
    For X = 1 To NumStudents
        Put #1, X, Student(X)
    Next X
    
    Close #1
    
    Exit Sub

'Skips the line that caused the error
ErrorHandler:
    Resume Next
End Sub

'Retrieves the file from the user for the purpose of opening it
Public Function GetFile(ByVal dialogBox As Control) As String
    
    With dialogBox
        .FileName = ""
        .Filter = "Text Files|*.txt|Record Files|*.rec"
        .InitDir = App.Path
        .ShowOpen
        
        GetFile = .FileName
    End With

End Function

'Retrieves the file from the user for the purpose of saving the contents of the student record
Public Function GetSaveFile(ByVal dialogBox As Control) As String

    With dialogBox
        .FileName = ""
        .Filter = "Record Files|*.rec"
        .InitDir = App.Path
        .ShowSave
        
        GetSaveFile = .FileName
    End With

End Function

'Deletes the selected record in the list box
Public Sub DeleteRecord(ByRef lstBox As Control, ByRef Stu() As StudentRec, ByRef NumStu As Integer)
    
    Dim K As Integer
    
    'Decrements total students
    NumStu = NumStu - 1
    
    'Shifts all the values of the list upward and changes the numbers
    For K = lstBox.ListIndex To NumStu
        If K <> NumStu Then
            Stu(K + 1) = Stu(K + 2)
            If VBA.Len(lstBox.List(K + 1)) > 4 Then
                lstBox.List(K) = VBA.Format$(K + 1, "@@@") & "." & VBA.Right$(lstBox.List(K + 1), VBA.Len(lstBox.List(K + 1)) - 4)
            Else
                lstBox.List(K) = lstBox.List(K + 1)
            End If
        End If
    Next K
    
    'Resets the list index
    lstBox.ListIndex = -1
    
    'Removes the last 'empty' item from the list box
    lstBox.RemoveItem lstBox.ListCount - 1
    
End Sub

'Validates each field in the changer form
Public Function Validate(ByRef NewStudent As StudentRec, _
                         ByRef textBox As Control) As Boolean
    
    Dim ErrorMsg As String
    Dim ErrorTitle As String
    Dim ErrorType As Integer
    
    Dim Valid As Boolean
    
    'Valid is initially true
    Valid = True
    ErrorType = vbExclamation + vbOKOnly
            
    With NewStudent
        'Checks to see if any of the fields are empty
        If VBA.Trim$(.LastName) = "" Then
            MsgBox "Please enter a lastname in the LastName field.", ErrorType, "Error: No LastName"
            Valid = False
        ElseIf VBA.Trim$(.FirstName) = "" Then
            MsgBox "Please enter a firstname in the FirstName field.", ErrorType, "Error: No FirstName"
            Valid = False
        ElseIf VBA.Trim$(.HomeForm) = "" Then
            MsgBox "Please enter a homeform in the HomeForm field.", ErrorType, "Error: No HomeForm"
            Valid = False
        ElseIf textBox.Text = "" Then
            MsgBox "Please enter a mark in the Mark field.", ErrorType, "Error: No Mark"
            Valid = False
        Else
            'Checks if the homeform is valid (the grade is 9 to 12)
            If (Val(.HomeForm) < 9 Or Val(.HomeForm) > 12) Or (VBA.Right$(VBA.Trim$(.HomeForm), 1) < "A" Or VBA.Right$(VBA.Trim$(.HomeForm), 1) > "Z") Then
                MsgBox "The HomeForm entered is not valid!", ErrorType, "Error: Invalid HomeForm"
                Valid = False
            'Checks if the mark is a valid number
            ElseIf .Mark < 0 Or .Mark > 100 Then
                MsgBox "Number exceeds the range of a valid mark! (0 to 100)", ErrorType, "Error: Invalid Mark"
                Valid = False
            End If
        End If
    End With
    
    'Assigns the functions value of Valid
    Validate = Valid

End Function
