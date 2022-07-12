VERSION 5.00
Begin VB.Form addnurses 
   Caption         =   "Add Nurse"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15390
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   15390
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Main"
      Height          =   10815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   20175
      Begin VB.Frame Frame2 
         BackColor       =   &H0000C0C0&
         Caption         =   "NURSE INFORMATION"
         Height          =   7815
         Left            =   720
         TabIndex        =   10
         Top             =   600
         Width           =   18735
         Begin VB.PictureBox Picture2 
            Height          =   4575
            Left            =   480
            Picture         =   "addnurse.frx":0000
            ScaleHeight     =   4515
            ScaleWidth      =   4515
            TabIndex        =   30
            Top             =   3120
            Width           =   4575
         End
         Begin VB.PictureBox Picture1 
            Height          =   4575
            Left            =   7080
            Picture         =   "addnurse.frx":3809
            ScaleHeight     =   4515
            ScaleWidth      =   7515
            TabIndex        =   29
            Top             =   3240
            Width           =   7575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "addnurse.frx":118D6
            Left            =   2280
            List            =   "addnurse.frx":118E0
            TabIndex        =   18
            Text            =   "Male"
            Top             =   1680
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   2280
            TabIndex        =   17
            Top             =   360
            Width           =   4695
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   2280
            TabIndex        =   16
            Top             =   960
            Width           =   4695
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   9720
            TabIndex        =   15
            Top             =   360
            Width           =   4695
         End
         Begin VB.TextBox Text8 
            Height          =   375
            Left            =   9720
            TabIndex        =   14
            Top             =   1080
            Width           =   4695
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   9720
            TabIndex        =   13
            Top             =   1800
            Width           =   4695
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   9720
            TabIndex        =   12
            Top             =   2520
            Width           =   4695
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   2280
            TabIndex        =   11
            Top             =   2400
            Width           =   4695
         End
         Begin VB.Label Label1 
            Caption         =   "Fathers Name"
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   26
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Nurse Salary"
            Height          =   375
            Index           =   5
            Left            =   7560
            TabIndex        =   25
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Nurse Department"
            Height          =   375
            Index           =   4
            Left            =   7560
            TabIndex        =   24
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Date of Birth"
            Height          =   375
            Index           =   3
            Left            =   7680
            TabIndex        =   23
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "NurseEduqualification"
            Height          =   375
            Index           =   2
            Left            =   7680
            TabIndex        =   22
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Nurse Name"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Nurse Id"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Gender"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   19
            Top             =   1680
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0000C0C0&
         Caption         =   "Control Buttons"
         Height          =   1455
         Left            =   720
         TabIndex        =   1
         Top             =   8760
         Width           =   18855
         Begin VB.CommandButton Command2 
            Height          =   495
            Left            =   1800
            Picture         =   "addnurse.frx":118F2
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Left            =   3240
            Picture         =   "addnurse.frx":1719D
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd2 
            Height          =   495
            Index           =   0
            Left            =   14640
            Picture         =   "addnurse.frx":1CA69
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd9 
            Height          =   495
            Index           =   1
            Left            =   4560
            Picture         =   "addnurse.frx":22320
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd5 
            Height          =   495
            Index           =   2
            Left            =   8880
            Picture         =   "addnurse.frx":27BA7
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd4 
            Height          =   495
            Index           =   3
            Left            =   7440
            Picture         =   "addnurse.frx":2D431
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd3 
            Height          =   495
            Index           =   4
            Left            =   6000
            Picture         =   "addnurse.frx":32CDC
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd6 
            Height          =   495
            Index           =   5
            Left            =   10320
            Picture         =   "addnurse.frx":38566
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd10 
            Height          =   495
            Index           =   6
            Left            =   13200
            Picture         =   "addnurse.frx":3DE43
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd8 
            Height          =   495
            Index           =   8
            Left            =   11760
            Picture         =   "addnurse.frx":4370A
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   480
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "addnurses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub nurse()
Set rs = New ADODB.Recordset
        rs.CursorType = adOpenDynamic
        rs.LockType = adLockOptimistic
        rs.Open "nurse", cn, , , adCmdTable
        Set rk = rs
        Set Text2.DataSource = rk
Text2.DataField = "nur_name"
Set Text4.DataSource = rk
Text4.DataField = "qualification"
Set Text8.DataSource = rk
Text8.DataField = "d_o_b"
Set Text6.DataSource = rk
Text6.DataField = "salary"
Set Text7.DataSource = rk
Text7.DataField = "fathers_name"
Set Text1.DataSource = rk
Text1.DataField = "nur_id"
Set Text5.DataSource = rk
Text5.DataField = "dept_id"
Set Combo1.DataSource = rk
Combo1.DataField = "dept_id"

End Sub


Private Sub cmd10_Click(Index As Integer)
Dim nurseno As Integer
nurseno = 1
sqk = "SELECT nur_name FROM (SELECT nur_name FROM nurse ORDER BY nur_name desc) WHERE ROWNUM <= 1 "

Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open sqk, cn, , , adCmdText
nurseno = rk.Fields(0)
Text1.Text = nurseno + 1

         
 rk.Close
        
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
On Error GoTo errordesc
rk.MoveFirst
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call nurse
End Sub

Private Sub cmd4_Click(Index As Integer)
On Error GoTo errordesc
If rk.EOF = True Then
    rk.MoveFirst
    
Else
rk.MoveNext
End If
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call nurse
End Sub

Private Sub cmd5_Click(Index As Integer)
On Error GoTo errordesc
rk.MoveLast
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call nurse
End Sub

Private Sub cmd6_Click(Index As Integer)
On Error GoTo errordesc
If (rk.BOF = True) Then
  rk.MoveLast
Else
  rk.MovePrevious
End If
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call nurse
End Sub

Private Sub cmd8_Click(Index As Integer)
Dim aid As Integer
aid = InputBox("Enter the doctor id to search", "search", "id")
sqk = "select * from nurse where nur_id= " & Val(aid)
Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open sqk, cn, , , adCmdText
If rk.EOF = True Then
MsgBox " nurse id not exist!!......"
Call nurse
End If
Set Text2.DataSource = rk
Text2.DataField = "nur_name"
Set Text4.DataSource = rk
Text4.DataField = "qualification"
Set Text8.DataSource = rk
Text8.DataField = "d_o_b"
Set Text6.DataSource = rk
Text6.DataField = "salary"
Set Text7.DataSource = rk
Text7.DataField = "fathers_name"

Set Text1.DataSource = rk
Text1.DataField = "nur_id"
Set Text5.DataSource = rk
Text5.DataField = "dept_id"
Set Combo1.DataSource = rk
Combo1.DataField = "dept_id"


End Sub

Private Sub cmd9_Click(Index As Integer)
rk.AddNew

End Sub

Private Sub Command1_Click()
rk.Delete

End Sub

Private Sub Command2_Click()
rk.Update
MsgBox "UPDATION OD DATA COMPLETED"
rk.MoveFirst
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
cn.CursorLocation = adUseClient
cn.Open

Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open "NURSE", cn, , , adCmdTable
MsgBox "loading please wait !!.............. ", , "Loading Message......."


Set Text2.DataSource = rk
Text2.DataField = "nur_name"
Set Text4.DataSource = rk
Text4.DataField = "qualification"
Set Text8.DataSource = rk
Text8.DataField = "d_o_b"
Set Text6.DataSource = rk
Text6.DataField = "salary"
Set Text7.DataSource = rk
Text7.DataField = "fathers_name"
Set Text1.DataSource = rk
Text1.DataField = "nur_id"
Set Text5.DataSource = rk
Text5.DataField = "dept_id"
Set Combo1.DataSource = rk
Combo1.DataField = "dept_id"

rk.MoveFirst



End Sub





