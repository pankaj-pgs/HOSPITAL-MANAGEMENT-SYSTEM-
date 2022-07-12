VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   8415
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   9645
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4971.86
   ScaleMode       =   0  'User
   ScaleWidth      =   9056.133
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   3120
      TabIndex        =   2
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   3120
      TabIndex        =   1
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "      LOGIN"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "   Hospital Management System"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   8475
      Left            =   0
      Picture         =   "frmLogin1.frx":0000
      Top             =   0
      Width           =   9630
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
End Sub



Private Sub Label2_Click()
If Text1.Text = "" And Text2.Text = "" Then
        LoginSucceeded = True
        
        Me.Hide
        frmSplash.Show
        
        
        
        
        
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        Text2.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub
