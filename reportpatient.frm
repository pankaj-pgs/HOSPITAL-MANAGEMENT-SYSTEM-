VERSION 5.00
Begin VB.Form reportpatient 
   Caption         =   "Patient Report"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   16020
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   7800
      Width           =   1215
   End
End
Attribute VB_Name = "reportpatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload reportpatient
End Sub
