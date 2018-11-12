Attribute VB_Name = "Module1"
Option Explicit

'Declares the student record
Type StudentRec
    LastName As String * 20
    FirstName As String * 15
    HomeForm As String * 3
    Mark As Integer
End Type

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
Public Sub DisplayAll(ByVal picBox As Control, ByVal lblTotal As Control, ByRef Student() As StudentRec, ByVal NumStudents As Integer)

    Dim X As Integer
    
    'Clears the picture box of previous content
    picBox.Cls
    
    'Prints the header and dash-line seperator
    picBox.Print "  #."; Tab(6); _
                 "LastName"; Tab(27); _
                 "FirstName"; Tab(44); _
                 "HF"; Tab(54); _
                 "Mark(%)"
    
    picBox.Print VBA.String$(60, "-")
    
    'Prints each field of the student record of the array
    For X = 1 To NumStudents
        picBox.Print " "; VBA.Format$(X, "@@"); ". "; _
                     Student(X).LastName; " "; _
                     Student(X).FirstName; " "; _
                     VBA.Format$(VBA.Trim$(Student(X).HomeForm), "@@@"); Tab(55); _
                     VBA.Format$(Student(X).Mark, "@@@")
    Next X
    
    'Displays the actual number of records
    lblTotal.Caption = VBA.Str$(NumStudents)
    
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
'with that file name
Public Function GetSaveFile(ByVal dialogBox As Control) As String

    With dialogBox
        .FileName = ""
        .Filter = "Record Files|*.rec"
        .InitDir = App.Path
        .ShowSave
        
        GetSaveFile = .FileName
    End With

End Function

