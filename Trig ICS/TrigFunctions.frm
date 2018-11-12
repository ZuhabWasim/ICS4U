VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trigonometric Functions"
   ClientHeight    =   7680
   ClientLeft      =   660
   ClientTop       =   2220
   ClientWidth     =   15420
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8224.899
   ScaleMode       =   0  'User
   ScaleWidth      =   17370.97
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Equation"
      Height          =   7575
      Left            =   10560
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset"
         Height          =   615
         Left            =   2640
         TabIndex        =   61
         Top             =   6840
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Range"
         Height          =   1215
         Left            =   120
         TabIndex        =   51
         Top             =   4560
         Width           =   4575
         Begin VB.Frame Frame10 
            Height          =   855
            Left            =   1080
            TabIndex        =   57
            Top             =   240
            Width           =   975
            Begin VB.OptionButton optYLess1 
               Caption         =   "<"
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   59
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton optYLessEqual1 
               Caption         =   "<="
               Height          =   330
               Left            =   120
               TabIndex        =   58
               Top             =   480
               Width           =   735
            End
         End
         Begin VB.TextBox txtYMin 
            Alignment       =   1  'Right Justify
            Height          =   495
            Left            =   120
            TabIndex        =   56
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtYMax 
            Alignment       =   1  'Right Justify
            Height          =   495
            Left            =   3600
            TabIndex        =   55
            Top             =   480
            Width           =   855
         End
         Begin VB.Frame Frame9 
            Height          =   855
            Left            =   2520
            TabIndex        =   52
            Top             =   240
            Width           =   975
            Begin VB.OptionButton optYLessEqual2 
               Caption         =   "<="
               Height          =   285
               Left            =   120
               TabIndex        =   54
               Top             =   480
               Width           =   735
            End
            Begin VB.OptionButton optYLess2 
               Caption         =   "<"
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   53
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            Caption         =   "y"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2040
            TabIndex        =   60
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Domain"
         Height          =   1215
         Left            =   120
         TabIndex        =   35
         Top             =   3240
         Width           =   4575
         Begin VB.Frame Frame8 
            Height          =   855
            Left            =   2520
            TabIndex        =   48
            Top             =   240
            Width           =   975
            Begin VB.OptionButton optXLess2 
               Caption         =   "<"
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   50
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton optXLessEqual2 
               Caption         =   "<="
               Height          =   285
               Left            =   120
               TabIndex        =   49
               Top             =   480
               Width           =   735
            End
         End
         Begin VB.TextBox txtXMax 
            Alignment       =   1  'Right Justify
            Height          =   495
            Left            =   3600
            TabIndex        =   47
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtXMin 
            Alignment       =   1  'Right Justify
            Height          =   495
            Left            =   120
            TabIndex        =   46
            Top             =   480
            Width           =   855
         End
         Begin VB.Frame Frame4 
            Height          =   855
            Left            =   1080
            TabIndex        =   43
            Top             =   240
            Width           =   975
            Begin VB.OptionButton optXLessEqual1 
               Caption         =   "<="
               Height          =   330
               Left            =   120
               TabIndex        =   45
               Top             =   480
               Width           =   735
            End
            Begin VB.OptionButton optXLess1 
               Caption         =   "<"
               Enabled         =   0   'False
               Height          =   330
               Left            =   120
               TabIndex        =   44
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2040
            TabIndex        =   42
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   615
         Left            =   1560
         TabIndex        =   24
         Top             =   6840
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   615
         Left            =   3720
         TabIndex        =   21
         Top             =   6840
         Width           =   975
      End
      Begin VB.CommandButton cmdGraph 
         Caption         =   "Graph It"
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Frame Frame7 
         Caption         =   "Type"
         Height          =   855
         Left            =   120
         TabIndex        =   16
         Top             =   5880
         Width           =   4575
         Begin VB.OptionButton optSin 
            Caption         =   "Sine"
            Height          =   330
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optCos 
            Caption         =   "Cosine"
            Height          =   330
            Left            =   1440
            TabIndex        =   18
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optTan 
            Caption         =   "Tangent"
            Height          =   375
            Left            =   2880
            TabIndex        =   17
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Shift"
         Height          =   2295
         Left            =   2520
         TabIndex        =   9
         Top             =   840
         Width           =   2175
         Begin VB.TextBox txtHShift 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   720
            TabIndex        =   11
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtVShift 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   720
            TabIndex        =   10
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "d ="
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "c ="
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Horizontal:"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "Vertical:"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Dilation"
         Height          =   2295
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2295
         Begin VB.TextBox txtVStretch 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   720
            TabIndex        =   4
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtHStretch 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   720
            TabIndex        =   3
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Vertical:"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label12 
            Caption         =   "Horizontal:"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "a ="
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "k ="
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   1680
            Width           =   615
         End
      End
      Begin VB.Label Label1 
         Caption         =   "f(x) = a[k(x - d)] + c"
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "-3"
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
      Left            =   4920
      TabIndex        =   41
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "-4"
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
      Left            =   4920
      TabIndex        =   40
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
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
      Left            =   5040
      TabIndex        =   39
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
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
      Left            =   5040
      TabIndex        =   38
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
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
      Left            =   5040
      TabIndex        =   37
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "-2"
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
      Left            =   4920
      TabIndex        =   36
      Top             =   5520
      Width           =   255
   End
   Begin VB.Line Line20 
      X1              =   5806.095
      X2              =   6206.01
      Y1              =   6999.732
      Y2              =   6999.732
   End
   Begin VB.Line Line19 
      X1              =   5806.095
      X2              =   6206.01
      Y1              =   7967.872
      Y2              =   7967.872
   End
   Begin VB.Line Line18 
      X1              =   5806.095
      X2              =   6206.01
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line17 
      X1              =   5806.095
      X2              =   6206.01
      Y1              =   1000.268
      Y2              =   1000.268
   End
   Begin VB.Line Line16 
      X1              =   5806.095
      X2              =   6206.01
      Y1              =   5000.268
      Y2              =   5000.268
   End
   Begin VB.Line Line15 
      X1              =   5806.095
      X2              =   6206.01
      Y1              =   2999.732
      Y2              =   2999.732
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "-900"
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
      Left            =   600
      TabIndex        =   34
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "-720"
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
      Left            =   1560
      TabIndex        =   33
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "-540"
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
      Left            =   2520
      TabIndex        =   32
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "-360"
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
      Left            =   3360
      TabIndex        =   31
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "-180"
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
      Left            =   4200
      TabIndex        =   30
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "900"
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
      Left            =   9720
      TabIndex        =   29
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "720"
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
      Left            =   8880
      TabIndex        =   28
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "540"
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
      Left            =   7920
      TabIndex        =   27
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "360"
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
      Left            =   7080
      TabIndex        =   26
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "180"
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
      Left            =   6120
      TabIndex        =   25
      Top             =   3840
      Width           =   375
   End
   Begin VB.Line Line14 
      X1              =   1000.352
      X2              =   1000.352
      Y1              =   3855.422
      Y2              =   4112.45
   End
   Begin VB.Line Line13 
      X1              =   4999.505
      X2              =   4999.505
      Y1              =   3855.422
      Y2              =   4112.45
   End
   Begin VB.Line Line12 
      X1              =   4000.28
      X2              =   4000.28
      Y1              =   3855.422
      Y2              =   4112.45
   End
   Begin VB.Line Line11 
      X1              =   2999.928
      X2              =   2999.928
      Y1              =   3855.422
      Y2              =   4112.45
   End
   Begin VB.Line Line10 
      X1              =   1999.577
      X2              =   1999.577
      Y1              =   3855.422
      Y2              =   4112.45
   End
   Begin VB.Line Line9 
      X1              =   10000.14
      X2              =   10000.14
      Y1              =   3855.422
      Y2              =   4112.45
   End
   Begin VB.Line Line8 
      X1              =   11000.49
      X2              =   11000.49
      Y1              =   3855.422
      Y2              =   4112.45
   End
   Begin VB.Line Line7 
      X1              =   8999.785
      X2              =   8999.785
      Y1              =   3855.422
      Y2              =   4112.45
   End
   Begin VB.Line Line6 
      X1              =   8000.56
      X2              =   8000.56
      Y1              =   3855.422
      Y2              =   4112.45
   End
   Begin VB.Line Line3 
      X1              =   7000.208
      X2              =   7000.208
      Y1              =   3855.422
      Y2              =   4112.45
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Left            =   5040
      TabIndex        =   23
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "-1"
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
      Left            =   4920
      TabIndex        =   22
      Top             =   4560
      Width           =   255
   End
   Begin VB.Line Line5 
      X1              =   5800.462
      X2              =   6200.377
      Y1              =   5999.465
      Y2              =   5999.465
   End
   Begin VB.Line Line4 
      X1              =   5800.462
      X2              =   6200.377
      Y1              =   2008.032
      Y2              =   2008.032
   End
   Begin VB.Line Line2 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   5999.857
      X2              =   5999.857
      Y1              =   0
      Y2              =   8000
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   0
      X2              =   11999.71
      Y1              =   4000
      Y2              =   4000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const XAXIS = 4000
Const YAXIS = 6000

Const XSTRETCH = 5.55555555555556
Const YSTRETCH = 1000


Private Sub cmdClear_Click()
    
    Cls
    
End Sub

Private Sub cmdExit_Click()
    
    If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Exit") = vbYes Then
        End
    End If
    
End Sub

Private Sub cmdGraph_Click()
    
    'Dilations
    Dim A As Single
    Dim K As Single
    
    'Shifts
    Dim C As Single
    Dim D As Single
    
    'Domain
    Dim XMin As Integer
    Dim XMax As Integer
    
    'Range
    Dim YMin As Integer
    Dim YMax As Integer
    
    'Gets the tranformation values inputed
    GetValues A, K, C, D, XMin, XMax, YMin, YMax
    
    'Graphs the function
    Graph XAXIS, YAXIS, XSTRETCH, YSTRETCH, A, K, C, D, XMin, XMax, YMin, YMax
    
    'Displays the values of each field in the given textbox
    DisplayValues A, K, C, D, XMin, XMax, YMin, YMax
    
End Sub

Private Sub cmdReset_Click()
    
    Reset
    
End Sub

Private Sub Form_Load()
    
    Reset
    
End Sub
