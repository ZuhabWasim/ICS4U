VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About"
   ClientHeight    =   2295
   ClientLeft      =   3090
   ClientTop       =   4365
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "A5_WasimZAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Return"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Random Access File Editor"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   120
      Top             =   480
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   5040
      Picture         =   "A5_WasimZAbout.frx":0442
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label Label3 
      Caption         =   "Version: 2.0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "© Zuhab Wasim 2017, All Rights Reserved"
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
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   $"A5_WasimZAbout.frx":0884
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Returns the user back to the main form
'by unloading the form
Private Sub cmdReturn_Click()
    
    Unload frmAbout
    
End Sub
