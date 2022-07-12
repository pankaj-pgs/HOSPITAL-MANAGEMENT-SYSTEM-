VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   9015
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5326.359
   ScaleMode       =   0  'User
   ScaleWidth      =   15098.25
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtPassword 
      BorderStyle     =   0  'None
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   8985
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4260
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10800
      Picture         =   "frmLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   1380
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   9000
      Picture         =   "frmLogin.frx":59E6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1380
   End
   Begin VB.TextBox txtUserName 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   8985
      TabIndex        =   0
      Top             =   3840
      Width           =   2325
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password Please Contact  Administrator..... "
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Top             =   6240
      Width           =   5775
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   7680
      TabIndex        =   5
      Top             =   4245
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   7680
      TabIndex        =   4
      Top             =   3855
      Width           =   1080
   End
   Begin VB.Image Image2 
      Height          =   3885
      Left            =   5400
      Picture         =   "frmLogin.frx":B39C
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   9915
   End
   Begin VB.Image Image1 
      Height          =   12600
      Left            =   0
      Picture         =   "frmLogin.frx":EA4A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20400
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdcancel_Click()
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If txtPassword = "" And txtUserName = "" Then
        LoginSucceeded = True
        
        Me.Hide
        frmSplash.Show
        
        
        
        
        
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

