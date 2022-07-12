VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4665
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   6840
      Top             =   3720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital  Management System"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
ProgressBar1.Value = ProgressBar1.Min

End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 10

If ProgressBar1.Value >= ProgressBar1.Max Then
Timer1.Enabled = False
End If
If ProgressBar1.Value = 100 Then
frmSplash.Hide
frmmain.Show
End If




End Sub

