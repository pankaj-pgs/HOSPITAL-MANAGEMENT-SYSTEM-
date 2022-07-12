VERSION 5.00
Begin VB.Form reportmedecine 
   Caption         =   " Medicine Report"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15990
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   15990
   Begin VB.CommandButton Command1 
      Caption         =   "back"
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   8400
      Width           =   1335
   End
End
Attribute VB_Name = "reportmedecine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload reportmedicine
End Sub
