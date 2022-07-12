VERSION 5.00
Begin VB.Form addaccount 
   Caption         =   "Add /Update/Remove   Account"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16125
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Main"
      Height          =   10815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   20175
      Begin VB.Frame Frame2 
         Caption         =   "Doctor information"
         Height          =   8415
         Left            =   720
         TabIndex        =   10
         Top             =   600
         Width           =   18855
         Begin VB.PictureBox Picture1 
            Height          =   2295
            Left            =   14040
            ScaleHeight     =   2235
            ScaleWidth      =   2475
            TabIndex        =   23
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox Text9 
            Height          =   375
            Index           =   5
            Left            =   2280
            TabIndex        =   22
            Top             =   5520
            Width           =   4815
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   4
            Left            =   2280
            TabIndex        =   21
            Top             =   960
            Width           =   4815
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   3
            Left            =   2280
            TabIndex        =   20
            Top             =   2400
            Width           =   4815
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   0
            Left            =   2280
            TabIndex        =   19
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   18
            Top             =   3720
            Width           =   4815
         End
         Begin VB.TextBox Text8 
            Height          =   375
            Index           =   6
            Left            =   2280
            TabIndex        =   17
            Top             =   4920
            Width           =   4815
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Index           =   7
            Left            =   2280
            TabIndex        =   16
            Top             =   4320
            Width           =   4815
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "addaccount.frx":0000
            Left            =   2280
            List            =   "addaccount.frx":000A
            TabIndex        =   15
            Text            =   "Male"
            Top             =   1680
            Width           =   2895
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "addaccount.frx":001C
            Left            =   2280
            List            =   "addaccount.frx":0038
            TabIndex        =   14
            Top             =   3000
            Width           =   1215
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   3840
            TabIndex        =   13
            Top             =   3000
            Width           =   1335
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   5400
            TabIndex        =   12
            Top             =   3000
            Width           =   1335
         End
         Begin VB.TextBox Text9 
            Height          =   375
            Index           =   0
            Left            =   2280
            TabIndex        =   11
            Top             =   6120
            Width           =   4815
         End
         Begin VB.Label Label3 
            Caption         =   "abcd@example.com"
            Height          =   375
            Left            =   7440
            TabIndex        =   35
            Top             =   6120
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Patient Photograph "
            Height          =   495
            Left            =   14400
            TabIndex        =   34
            Top             =   3600
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Phone Number"
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   33
            Top             =   5520
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Fathers Name"
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   32
            Top             =   4920
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Salary"
            Height          =   375
            Index           =   5
            Left            =   360
            TabIndex        =   31
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Department"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   30
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Date of Birth"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   29
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Eduqualification"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   28
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Name"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Id"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Gender"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   25
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail id"
            Height          =   375
            Index           =   9
            Left            =   360
            TabIndex        =   24
            Top             =   6240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Control Buttons"
         Height          =   1335
         Left            =   720
         TabIndex        =   1
         Top             =   9240
         Width           =   19095
         Begin VB.CommandButton cmd2 
            Caption         =   "BACK"
            Height          =   375
            Index           =   0
            Left            =   14640
            TabIndex        =   9
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd2 
            Caption         =   "SAVE"
            Height          =   375
            Index           =   1
            Left            =   4680
            TabIndex        =   8
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd5 
            Caption         =   "LAST"
            Height          =   375
            Index           =   2
            Left            =   8880
            TabIndex        =   7
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd4 
            Caption         =   "NEXT"
            Height          =   375
            Index           =   3
            Left            =   7440
            TabIndex        =   6
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd3 
            Caption         =   "FIRST"
            Height          =   375
            Index           =   4
            Left            =   6000
            TabIndex        =   5
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd6 
            Caption         =   "PREVIOUS"
            Height          =   375
            Index           =   5
            Left            =   10320
            TabIndex        =   4
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd10 
            Caption         =   "EXIT"
            Height          =   375
            Index           =   6
            Left            =   13200
            TabIndex        =   3
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd8 
            Caption         =   "FIND"
            Height          =   375
            Index           =   8
            Left            =   11760
            TabIndex        =   2
            Top             =   600
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "addaccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
