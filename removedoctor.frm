VERSION 5.00
Begin VB.Form removedoctor 
   Caption         =   "Update/Remove Doctor"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16155
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   16155
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Main"
      Height          =   10815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20175
      Begin VB.Frame Frame2 
         BackColor       =   &H0000C0C0&
         Caption         =   "Doctor information"
         Height          =   8415
         Left            =   480
         TabIndex        =   10
         Top             =   600
         Width           =   19335
         Begin VB.PictureBox Picture1 
            Height          =   6015
            Left            =   7920
            Picture         =   "removedoctor.frx":0000
            ScaleHeight     =   5955
            ScaleWidth      =   8595
            TabIndex        =   31
            Top             =   240
            Width           =   8655
         End
         Begin VB.TextBox Text9 
            Height          =   375
            Left            =   2280
            TabIndex        =   30
            Top             =   6120
            Width           =   4455
         End
         Begin VB.TextBox Text8 
            Height          =   375
            Left            =   2280
            TabIndex        =   29
            Top             =   5520
            Width           =   4455
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   2280
            TabIndex        =   28
            Top             =   4920
            Width           =   4455
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   2280
            TabIndex        =   27
            Top             =   4320
            Width           =   4455
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   2280
            TabIndex        =   26
            Top             =   3720
            Width           =   4455
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   2280
            TabIndex        =   25
            Top             =   2280
            Width           =   4455
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   2280
            TabIndex        =   24
            Top             =   3000
            Width           =   4455
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   2280
            TabIndex        =   23
            Top             =   960
            Width           =   4455
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   2280
            TabIndex        =   22
            Top             =   360
            Width           =   2175
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "removedoctor.frx":11B97
            Left            =   2280
            List            =   "removedoctor.frx":11BA1
            TabIndex        =   11
            Text            =   "Male"
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Phone Number"
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   21
            Top             =   5520
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Fathers Name"
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   20
            Top             =   4920
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Patient Salary"
            Height          =   375
            Index           =   5
            Left            =   360
            TabIndex        =   19
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Department"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   18
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Date of Birth"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   17
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Eduqualification"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   16
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Name"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   15
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000A&
            Caption         =   "Doctor Id"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Gender"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   13
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail id"
            Height          =   375
            Index           =   9
            Left            =   480
            TabIndex        =   12
            Top             =   6120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0000C0C0&
         Caption         =   "Control Buttons"
         Height          =   1335
         Left            =   480
         TabIndex        =   1
         Top             =   9240
         Width           =   19335
         Begin VB.CommandButton cmd2 
            Height          =   495
            Index           =   0
            Left            =   14640
            Picture         =   "removedoctor.frx":11BB3
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd9 
            Height          =   495
            Index           =   1
            Left            =   4680
            Picture         =   "removedoctor.frx":1746A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd5 
            Height          =   495
            Index           =   2
            Left            =   8880
            Picture         =   "removedoctor.frx":1CBF9
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd4 
            Height          =   495
            Index           =   3
            Left            =   7440
            Picture         =   "removedoctor.frx":22483
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd3 
            Height          =   495
            Index           =   4
            Left            =   6000
            Picture         =   "removedoctor.frx":27D2E
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd6 
            Height          =   495
            Index           =   5
            Left            =   10320
            Picture         =   "removedoctor.frx":2D5B8
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd10 
            Height          =   495
            Index           =   6
            Left            =   13200
            Picture         =   "removedoctor.frx":32E95
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd8 
            Height          =   495
            Index           =   8
            Left            =   11760
            Picture         =   "removedoctor.frx":3875C
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   480
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "removedoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmd10_Click(Index As Integer)
         Text1.Text = " "
         Text2.Text = " "
         Text3.Text = " "
         Text7.Text = " "
         Text8.Text = " "
         Text9.Text = " "
         Text10.Text = " "
         Text6.Text = " "
         Text5.Text = " "
         Text4.Text = " "
End Sub


Private Sub cmd2_Click(Index As Integer)
Unload Me


End Sub

Private Sub cmd3_Click(Index As Integer)
rk.MoveFirst

End Sub

Private Sub cmd4_Click(Index As Integer)
If rk.EOF = True Then
    rk.MoveFirst
Else
rk.MoveNext
End If


End Sub

Private Sub cmd5_Click(Index As Integer)
rk.MoveLast

End Sub

Private Sub cmd6_Click(Index As Integer)
If (rk.BOF = True) Then
 rk.MoveLast
Else
 rk.MovePrevious
End If
 
End Sub

Private Sub cmd8_Click(Index As Integer)
Dim aid As Integer
aid = InputBox("Enter the doctor id to search", "search", "id")
sqk = "select * from doctor where doctor_id=" & Val(aid)
Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open sqk, cn, , , adCmdText
Set Text2.DataSource = rk
Text2.DataField = "doct_name"
Set Text4.DataSource = rk
Text4.DataField = "qualification"
Set Text3.DataSource = rk
Text3.DataField = "d_o_b"
Set Text6.DataSource = rk
Text6.DataField = "salary"
Set Text7.DataSource = rk
Text7.DataField = "fathers_name"
Set Text8.DataSource = rk
Text8.DataField = "phn_no"
Set Text9.DataSource = rk
Text9.DataField = "email_id"
Set Text1.DataSource = rk
Text1.DataField = "doctor_id"
Set Text5.DataSource = rk
Text5.DataField = "dept_id"
Set Combo1.DataSource = rk
Combo1.DataField = "sex"
End Sub

Private Sub cmd9_Click(Index As Integer)
rk.Delete
rk.Update
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
cn.CursorLocation = adUseClient
cn.Open

Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open "doctor", cn, , , adCmdTable
MsgBox "loading please wait !!.............. ", , "Loading Message......."

Set Text2.DataSource = rk
Text2.DataField = "doct_name"
Set Text4.DataSource = rk
Text4.DataField = "qualification"
Set Text3.DataSource = rk
Text3.DataField = "d_o_b"
Set Text6.DataSource = rk
Text6.DataField = "salary"
Set Text7.DataSource = rk
Text7.DataField = "fathers_name"
Set Text8.DataSource = rk
Text8.DataField = "phn_no"
Set Text9.DataSource = rk
Text9.DataField = "email_id"
Set Text1.DataSource = rk
Text1.DataField = "doctor_id"
Set Text5.DataSource = rk
Text5.DataField = "dept_id"
Set Combo1.DataSource = rk
Combo1.DataField = "sex"

rk.MoveFirst



End Sub


