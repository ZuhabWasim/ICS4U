VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Toys! Toys! Toys!"
   ClientHeight    =   5850
   ClientLeft      =   1860
   ClientTop       =   2805
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "A2_WasimZ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   10935
   Begin MSComDlg.CommonDialog cdlDialog 
      Left            =   4440
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picChart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      ScaleHeight     =   4995
      ScaleWidth      =   10635
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   10695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   6
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdShowChart 
      Caption         =   "&Show Chart"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open..."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin VB.PictureBox picData 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      ScaleHeight     =   4995
      ScaleWidth      =   10635
      TabIndex        =   0
      Top             =   120
      Width           =   10695
   End
   Begin VB.Label lblTotalSales 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label lblTSalesDisplay 
      Caption         =   "Total Sales:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   5400
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Zuhab Wasim
'Date: 20/10/2016
'Purpose: To demonstrate and apply the knowledge the understanding of
'         2-dimensional arrays learned in the class of ICS4U

'Ensures variable declaration
Option Explicit

'Constants available to the entire form stated
Const TOY_MAX = 20
Const STORES_MAX = 4

'The arrays representing each toy's name and price declared
Dim ToyNames(1 To TOY_MAX) As String
Dim ToyPrice(1 To TOY_MAX) As Single

'The arrays representing the calculated sales amount and sold amount declared
Dim TotalSold(1 To TOY_MAX) As Integer
Dim TotalSales(1 To TOY_MAX) As Single

'The 2-dimensional array used for each array's store location declared
Dim StoreSales(1 To TOY_MAX, 1 To STORES_MAX) As Integer

'Variable representing the actual amount of toys in the file declared
Dim TotalToys As Integer

'Variable representing the combined total sales acheived from every toy declared
Dim CombinedSales As Single

'Variable that is used for the display of the chart declared
'Determines if the chart has already been printed on by with the same file
Dim IsPrinted As Boolean

'Code executed that returns the user back to the main display after viewing the sales chart
Private Sub cmdReturn_Click()
    
    'Toggles the visibility of each control so that the controls relating to the chart are invisibile
    ToggleChart picData, picChart, cmdOpen, cmdShowChart, cmdExit, cmdReturn, lblTSalesDisplay, lblTotalSales
    
    'Renames the form for the user to indentify that they are at the main display
    frmMain.Caption = "Toys! Toys! Toys! (Main)"
    
End Sub

'Code to be executed during the loading of the main form frmMain
Private Sub Form_Load()
    
    'Initially sets the visibility of return to false,
    'The ability to click on show chart denied,
    cmdReturn.Visible = False
    cmdShowChart.Enabled = False
    
End Sub

'Code to be executed when the user desires to open a file
Private Sub cmdOpen_Click()
    
    'Variable used for the file name declared
    Dim FileName As String
    
    'Retrieves a file to open from the user using window's common dialog control
    FileName = GetFile(cdlDialog)
    
    'Checks to see if the user has selected a text file to open
    If FileName <> "" And VBA.Right$(FileName, 3) = "txt" Then
        'Initializes all variables needed to be initialized
        Initialize ToyNames(), ToyPrice(), TOY_MAX, StoreSales(), STORES_MAX, TotalSold(), TotalSales(), CombinedSales
        'Sets IsPrinted to false to signify that the chart must display new content
        IsPrinted = False
        'Reads the file obtained from the user and stores it into the for variables
        ReadFile FileName, ToyNames(), ToyPrice(), TOY_MAX, StoreSales(), STORES_MAX, TotalToys
        'Calculates the needed values for total toys sold and the total and combined sales each toy has made
        Calculate ToyPrice(), TotalToys, StoreSales(), STORES_MAX, TotalSold(), TotalSales(), CombinedSales
        'Displays the contents of the file onto the picturebox picData
        DisplayData picData, lblTotalSales, ToyNames(), ToyPrice(), TotalToys, StoreSales(), STORES_MAX, TotalSold(), TotalSales(), CombinedSales
        'Allows the user to now view the chart
        cmdShowChart.Enabled = True
    End If
    
    

End Sub

'Code to be executed when the user desires to see the sales chart
Private Sub cmdShowChart_Click()
    
    'Toggles the visibility of each control so that the controls relating to the data are invisible
    ToggleChart picData, picChart, cmdOpen, cmdShowChart, cmdExit, cmdReturn, lblTSalesDisplay, lblTotalSales
    
    'Checks to see if the chart needs to print from a newly chosen file
    If IsPrinted = False Then
        'Displays the contents of the chart on to frmMain
        DisplayChart picChart, ToyNames(), TotalSold(), TOY_MAX
        'Assigns true to IsPrinted to notify that there are currently no need to redisplay any new content on the chart
        IsPrinted = True
    End If
    
    'Renames the form for the user to indentify that they are at the chart display
    frmMain.Caption = "Toys! Toys! Toys! (Sales Chart)"
    
End Sub

'Code to be executed when the user wants to exit the program
Private Sub cmdExit_Click()
        
    'Local exit variables declared
    Dim EMsg As String
    Dim EType As Integer
    Dim ETitle As String
    Dim EResponse As Integer

    'Exit variables assigned values
    EMsg = "Are you sure you want to exit?"
    EType = vbInformation + vbYesNo
    ETitle = "Exit"
    
    'Displays the messagebox and assigns the answer to EResponse
    EResponse = MsgBox(EMsg, EType, ETitle)

    'Exits the program if the user confirms their decision
    If EResponse = vbYes Then
        End
    End If
    
End Sub

