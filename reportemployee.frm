VERSION 5.00
Begin VB.Form reportemployee 
   Caption         =   "Report Employee"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   8400
      Width           =   1335
   End
End
Attribute VB_Name = "reportemployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload reportmedicine
End Sub
