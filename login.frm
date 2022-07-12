VERSION 5.00
Begin VB.Form login 
   Caption         =   "Form1"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16215
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   16215
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcancel 
      Height          =   495
      Index           =   1
      Left            =   10080
      Picture         =   "login.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdok 
      Height          =   495
      Index           =   0
      Left            =   8040
      Picture         =   "login.frx":59E6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox txtpassword 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   9600
      TabIndex        =   3
      Top             =   5520
      Width           =   2895
   End
   Begin VB.TextBox txtusername 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   2
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password Please Contact  Administrator..... "
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   7440
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   375
      Index           =   1
      Left            =   7800
      TabIndex        =   1
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   4965
      Left            =   5040
      Picture         =   "login.frx":B39C
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   10995
   End
   Begin VB.Image Image1 
      Height          =   12600
      Left            =   0
      Picture         =   "login.frx":EA4A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20400
   End
End
Attribute VB_Name = "login"
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

Private Sub cmdok_Click()
If (txtpassword = "" And txtusername = "") Then
        LoginSucceeded = True
        Me.Hide
        frmSplash.Show
        
        
        
        
        
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtpassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub
