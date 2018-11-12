Attribute VB_Name = "Module1"
Option Explicit

'Initializes and sets values for all variables needed to be initialized
Public Sub Initialize(ByRef TNames() As String, ByRef TPrice() As Single, ByVal TMAX As Integer, _
                      ByRef SSales() As Integer, ByVal SMAX As Integer, _
                      ByRef TSold() As Integer, ByRef TSales() As Single, ByRef CSales As Single)
    
    'Loop variables declared
    Dim R As Integer
    Dim C As Integer
    
    'Initializes all parameters
    CSales = 0
    For R = 1 To TMAX
        TNames(R) = ""
        TPrice(R) = 0
        TSold(R) = 0
        TSales(R) = 0
        For C = 1 To SMAX
            SSales(R, C) = 0
        Next C
    Next R
    
End Sub

'Using the common dialog box, returns the file the user has chosen to open
Public Function GetFile(ByVal cdlDialog As Control) As String
    
    'Asks for Text Files, opens in where the program is located, is used for opening a file
    With cdlDialog
        .FileName = ""
        .InitDir = App.Path
        .Filter = "Text Files|*.txt|All Files|*.*"
        .ShowOpen
        GetFile = .FileName
    End With
    
End Function

'Reads the file chosen by the user and stores the contents into the form variables of frmMain
Public Sub ReadFile(ByVal FName As String, _
                    ByRef TNames() As String, ByRef TPrice() As Single, ByVal TMAX As Integer, _
                    ByRef SSales() As Integer, ByVal SMAX As Integer, _
                    ByRef TToys As Integer)
    
    'Local variables declared
    Dim R As Integer
    Dim C As Integer
    
    'Initializes the counter variable
    R = 0
    
    'Opens the file for input
    Open FName For Input As #1
    
    'Reads the file and stores the toy names, toy prices into linear arrays
    'Stores the store sales into a bidimensional array
    Do While Not EOF(1)
        R = R + 1
        Input #1, TNames(R)
        Input #1, TPrice(R)
        For C = 1 To SMAX
            Input #1, SSales(R, C)
        Next C
    Loop
    
    'Closes the file
    Close #1
    
    'Assigns the actual amount of entries in the sequential file into the form variable TotalToys
    TToys = R

End Sub

'Calculates the total number of toys sold, as well as the amount of money earned
Public Sub Calculate(ByRef TPrice() As Single, ByVal TMAX As Integer, _
                     ByRef SSales() As Integer, ByVal SMAX As Integer, _
                     ByRef TSold() As Integer, ByRef TSales() As Single, ByRef CSales As Single)
    
    'Local loop and sum variables declared
    Dim TotalSum As Integer
    Dim R As Integer
    Dim C As Integer
    
    'Initializes sum variables that may or must be initialized
    TotalSum = 0
    CSales = 0
    
    'Adds all of the store sales and stores the total number of toys sold for each toy, total sales of each toy,
    'and the revenue from all toys combined
    For R = 1 To TMAX
        For C = 1 To SMAX
            TotalSum = TotalSum + SSales(R, C)
        Next C
        TSold(R) = TotalSum
        'Code that calculates the total amount of money earned from each toy
        TSales(R) = TotalSum * TPrice(R)
        CSales = CSales + TSales(R)
        'Re-assigns the sum variable TotalSum to zero for next iteration of the loop
        TotalSum = 0
    Next R
    
End Sub

'Displays the contents needed for the chart picture box
Public Sub DisplayChart(ByVal picChart As Control, _
                        ByRef TNames() As String, ByRef TSold() As Integer, ByVal TMAX As Integer)
    
    'Declares local variables
    Dim R As Integer
    Dim N As Integer
    'This variable represents the amount of cells (representing 5 or less sales) each toy has
    Dim NZeros As Integer
    
    'Clears the previous contents from the picture box
    picChart.Cls
    
    'Displays the header of the chart picture box
    picChart.Print " Toy Name"; Tab(22); "0"; Tab(32);
    For R = 1 To 7
        picChart.Print VBA.Format$((50 * R), "#"); Tab(42 + (10 * (R - 1)));
    Next R
    picChart.Print
    
    'Skips a line as the headers of the chart picture box do not take up 2 lines as the data picture box does
    picChart.Print
    
    'Displays a dashed line that seperates the headers from the content
    picChart.Print VBA.String$(101, "-")
    
    'Loops through each toy, printing its name and amount of zero's that represent the total toys sold
    For R = 1 To TMAX
        'Prints toy names
        'Also checks to see if the name is greater than 15 characters, cuts it off if otherwise
        If Len(TNames(R)) > 15 Then
            picChart.Print " "; VBA.Left$(TNames(R), 15) & "...";
        Else
            picChart.Print " "; TNames(R);
        End If
        'Checks to see the amount of toys evenly divides by 5
        'It will add an extra zero representing the remainder
        If TSold(R) Mod 5 = 0 Then
            NZeros = TSold(R) / 5
        Else
            NZeros = TSold(R) \ 5 + 1
        End If
        'Displays the zeros of each toy
        picChart.Print Tab(22); VBA.String$(NZeros, "0")
    Next R
    
End Sub

'Displays the contents of each array for the purpose of showing the data and sales of the variables
Public Sub DisplayData(ByVal picBox As Control, ByVal label As Control, _
                       ByRef TNames() As String, ByRef TPrice() As Single, ByVal TMAX As Integer, _
                       ByRef SSales() As Integer, ByVal SMAX As Integer, _
                       ByRef TSold() As Integer, ByRef TSales() As Single, ByVal CSales As Single)
    
    'Local loop variables delcared
    Dim R As Integer
    Dim C As Integer
    
    'Clears the picturebox for any previous prints
    picBox.Cls
    
    'Prints the first line of the headers ofor the data picture box
    picBox.Print " Toy"; Tab(37);
    For R = 1 To SMAX
        picBox.Print GetDirection(R); Tab(37 + (10 * R));
    Next R
    picBox.Print "Total"; Tab(91); " Toy"
    
    'Prints the second line of the headers for the data picture box
    picBox.Print " Description"; Tab(28); "Price"; Tab(37);
    For R = 1 To SMAX
        picBox.Print "Store"; Tab(37 + (10 * R));
    Next R
    picBox.Print " Sold"; Tab(91); "Sales"
    
    'Prints a dashed line to seperate the headers from the content of the picture box
    picBox.Print VBA.String$(101, "-")
    
    'Prints the contents of each toy into the picture box
    For R = 1 To TMAX
        'Prints the name of each toy
        'Also checks to see if the name is greater than 15 characters, cuts it off if otherwise
        If Len(TNames(R)) > 15 Then
            picBox.Print " "; VBA.Left$(TNames(R), 15) & "...";
        Else
            picBox.Print " "; TNames(R);
        End If
        picBox.Print Tab(25); "$ "; VBA.Format$(VBA.Format$(TPrice(R), "#,##0.00"), "@@@@@@@@"); Tab(37);
        'Prints the amount sold for each store of the toy and right aligns the values
        For C = 1 To SMAX
            picBox.Print VBA.Format$(SSales(R, C), "@@@@"); Tab(37 + (10 * C));
        Next C
        'Prints the total toys sold and total sales made of each toy and right aligns them
        picBox.Print VBA.Format$(TSold(R), "@@@@"); Tab(87); "$ "; VBA.Format$(VBA.Format$(TSales(R), "###,##0.00"), "@@@@@@@@@@@@")
    Next R
    
    'Prints the total number of sales made with every toy into a label on frmMain
    label.Caption = VBA.Format$(CSales, "currency")
    
End Sub

'Returns the direction given a value between 1 and 4, (1: " EAST", 2: "NORTH", 3: "SOUTH", 4: " WEST")
Public Function GetDirection(ByVal X As Integer) As String
    
    'Local variable declared
    Dim St As String
    
    'String used to make the substring of the directions needed
    St = " EASTNORTHSOUTH WEST"
    
    'Returns the substring needed given the value
    GetDirection = VBA.Mid$(St, 1 + 5 * (X - 1), 5)
    
End Function

'Toggles the visibility of components on the form for when the user wants to see the chart
'or when they want to see the data of each toy
Public Sub ToggleChart(ByVal picData As Control, ByVal picChart As Control, _
                       ByVal cmdOpen As Control, ByVal cmdShowChart As Control, ByVal cmdExit As Control, ByVal cmdReturn As Control, _
                       ByVal lblTSDisplay As Control, ByVal lblTSales As Control)
    
    'Toggles or reverses the visibility of each control
    'If its visible, make it invisible and vice versa
    picData.Visible = Not (picData.Visible)
    picChart.Visible = Not (picChart.Visible)
    
    cmdOpen.Visible = Not (cmdOpen.Visible)
    cmdShowChart.Visible = Not (cmdShowChart.Visible)
    cmdExit.Visible = Not (cmdExit.Visible)
    
    lblTSDisplay.Visible = Not (lblTSDisplay.Visible)
    lblTSales.Visible = Not (lblTSales.Visible)
    
    cmdReturn.Visible = Not (cmdReturn.Visible)
    
End Sub

