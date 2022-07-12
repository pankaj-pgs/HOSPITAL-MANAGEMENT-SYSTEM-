VERSION 5.00
Begin VB.Form reportdoctor 
   Caption         =   "Doctor Report"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16050
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   16050
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   8160
      Width           =   1335
   End
End
Attribute VB_Name = "reportdoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload reportdoctor
End Sub
