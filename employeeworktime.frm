VERSION 5.00
Begin VB.Form employeeworktime 
   Caption         =   "Employee Working"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16095
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   16095
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Main"
      Height          =   10815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   20175
      Begin VB.Frame Frame3 
         Caption         =   "Control Buttons"
         Height          =   1335
         Left            =   480
         TabIndex        =   16
         Top             =   9240
         Width           =   19335
         Begin VB.CommandButton cmd8 
            Caption         =   "FIND"
            Height          =   375
            Index           =   8
            Left            =   11760
            TabIndex        =   24
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd10 
            Caption         =   "EXIT"
            Height          =   375
            Index           =   6
            Left            =   13200
            TabIndex        =   23
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd6 
            Caption         =   "PREVIOUS"
            Height          =   375
            Index           =   5
            Left            =   10320
            TabIndex        =   22
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd3 
            Caption         =   "FIRST"
            Height          =   375
            Index           =   4
            Left            =   6000
            TabIndex        =   21
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd4 
            Caption         =   "NEXT"
            Height          =   375
            Index           =   3
            Left            =   7440
            TabIndex        =   20
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd5 
            Caption         =   "LAST"
            Height          =   375
            Index           =   2
            Left            =   8880
            TabIndex        =   19
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd2 
            Caption         =   "SAVE"
            Height          =   375
            Index           =   1
            Left            =   4680
            TabIndex        =   18
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd2 
            Caption         =   "BACK"
            Height          =   375
            Index           =   0
            Left            =   14640
            TabIndex        =   17
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Employee Working Schedule"
         Height          =   8415
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   19335
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   8
            Left            =   2280
            TabIndex        =   8
            Top             =   1680
            Width           =   4815
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   7
            Top             =   3720
            Width           =   4815
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   0
            Left            =   2280
            TabIndex        =   6
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Index           =   2
            Left            =   2280
            TabIndex        =   5
            Top             =   3120
            Width           =   4815
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   3
            Left            =   2280
            TabIndex        =   4
            Top             =   2400
            Width           =   4815
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   4
            Left            =   2280
            TabIndex        =   3
            Top             =   960
            Width           =   4815
         End
         Begin VB.PictureBox Picture1 
            Height          =   2295
            Left            =   14040
            ScaleHeight     =   2235
            ScaleWidth      =   2475
            TabIndex        =   2
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Designation"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   15
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Code"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   " Name"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Patient Eduqualification"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   12
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Time"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   11
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Day"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   10
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Patient Photograph "
            Height          =   495
            Left            =   14400
            TabIndex        =   9
            Top             =   3600
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "employeeworktime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd2_Click(index As Integer)
Unload employeeworktime
End Sub
