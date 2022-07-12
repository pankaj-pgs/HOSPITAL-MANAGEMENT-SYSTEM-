VERSION 5.00
Begin VB.Form departments 
   Caption         =   "Department"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16050
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8955
   ScaleWidth      =   16050
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
         Caption         =   "Department information"
         Height          =   7455
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   19335
         Begin VB.PictureBox Picture1 
            Height          =   6135
            Left            =   8760
            Picture         =   "department.frx":0000
            ScaleHeight     =   6075
            ScaleWidth      =   6075
            TabIndex        =   24
            Top             =   360
            Width           =   6135
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   2280
            TabIndex        =   23
            Top             =   3480
            Width           =   4815
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   2280
            TabIndex        =   22
            Top             =   2880
            Width           =   4815
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   2280
            TabIndex        =   21
            Top             =   2160
            Width           =   4815
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   2280
            TabIndex        =   20
            Top             =   1560
            Width           =   4815
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   2280
            TabIndex        =   19
            Top             =   960
            Width           =   4815
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   2280
            TabIndex        =   18
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Phone Number"
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   16
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Department Salary"
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   15
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Department Head"
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   14
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Department Doctor "
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   13
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Department Name"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Department Id"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0000C0C0&
         Caption         =   "Control Buttons"
         Height          =   1335
         Left            =   360
         TabIndex        =   1
         Top             =   8640
         Width           =   19335
         Begin VB.CommandButton Command1 
            Height          =   495
            Left            =   3240
            Picture         =   "department.frx":5FC9
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd2 
            Height          =   495
            Index           =   0
            Left            =   1800
            Picture         =   "department.frx":B758
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd7 
            Height          =   495
            Index           =   1
            Left            =   4680
            Picture         =   "department.frx":1100F
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd5 
            Height          =   495
            Index           =   2
            Left            =   8880
            Picture         =   "department.frx":16896
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd4 
            Height          =   495
            Index           =   3
            Left            =   7440
            Picture         =   "department.frx":1C120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd3 
            Height          =   495
            Index           =   4
            Left            =   6000
            Picture         =   "department.frx":219CB
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd9 
            Height          =   495
            Index           =   5
            Left            =   10320
            Picture         =   "department.frx":27255
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd10 
            Height          =   495
            Index           =   6
            Left            =   13200
            Picture         =   "department.frx":2CB32
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd8 
            Height          =   495
            Index           =   8
            Left            =   11760
            Picture         =   "department.frx":323F9
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   480
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "departments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub department()
Set rs = New ADODB.Recordset
        rs.CursorType = adOpenDynamic
        rs.LockType = adLockOptimistic
        rs.Open "department", cn, , , adCmdTable
        Set rk = rs
Set Text1.DataSource = rk
Text1.DataField = "dept_id"
Set Text2.DataSource = rk
Text2.DataField = "DEPT_NAME"
Set Text4.DataSource = rk
Text4.DataField = "DEPT_DOCTOR"
Set Text5.DataSource = rk
Text5.DataField = "DEPT_HEAD"
Set Text6.DataSource = rk
Text6.DataField = "DEPT_SALARY"
Set Text7.DataSource = rk
Text7.DataField = "PHN_NO"
        
End Sub



Private Sub cmd10_Click(Index As Integer)
Text1.Text = " "
         Text2.Text = " "
         Text7.Text = " "
         Text6.Text = " "
         Text5.Text = " "
         Text4.Text = " "
         Text1.Text = " "
End Sub



Private Sub cmd3_Click(Index As Integer)
On Error GoTo errordesc
rk.MoveFirst
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call department
End Sub

Private Sub cmd4_Click(Index As Integer)
On Error GoTo errordesc

If rk.EOF = True Then
    rk.MoveFirst
    
End If
rk.MoveNext
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call department
End Sub

Private Sub cmd5_Click(Index As Integer)
On Error GoTo errordesc

rk.MoveLast
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call department
End Sub

Private Sub cmd6_Click(Index As Integer)
On Error GoTo errordesc

rk.MoveLast
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call department
End Sub

Private Sub cmd7_Click(Index As Integer)
On Error GoTo errordesc
Set rk = New ADODB.Recordset
        rk.CursorType = adOpenDynamic
        rk.LockType = adLockOptimistic
        sqk = "select * from department "
        rk.Open sqk, cn, , , adCmdText
        rk.AddNew
        rk.Fields(0) = Text1.Text
        rk.Fields(1) = Text2.Text
        rk.Fields(2) = Text4.Text
        rk.Fields(3) = Text5.Text
        rk.Fields(4) = Text6.Text
        rk.Fields(5) = Text7.Text
        

        rk.Update
        Set Text1.DataSource = rk
Text1.DataField = "dept_id"
Set Text2.DataSource = rk
Text2.DataField = "DEPT_NAME"
Set Text3.DataSource = rk
Text3.DataField = "DEPT_SALARY"
Set Text4.DataSource = rk
Text4.DataField = "DEPT_DOCTOR"
Set Text5.DataSource = rk
Text5.DataField = "DEPT_HEAD"
Set Text6.DataSource = rk
Text6.DataField = "DEPT_SALARY"
Set Text7.DataSource = rk
Text7.DataField = "PHN_NO"


        
        rk.MoveFirst
        

Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call department
End Sub

Private Sub cmd8_Click(Index As Integer)
Dim aid As Integer
aid = InputBox("Enter the patientid to search", "search", "id")
sqk = "select * from DEPARTMENT where DEPt_id=" & Val(aid)
Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open sqk, cn, , , adCmdText

If rk.EOF = True Then
MsgBox "department id  not exist"
Call department
End If
Set Text1.DataSource = rk
Text1.DataField = "dept_id"
Set Text2.DataSource = rk
Text2.DataField = "DEPT_NAME"
Set Text4.DataSource = rk
Text4.DataField = "DEPT_DOCTOR"
Set Text5.DataSource = rk
Text5.DataField = "DEPT_HEAD"
Set Text6.DataSource = rk
Text6.DataField = "DEPT_SALARY"
Set Text7.DataSource = rk
Text7.DataField = "PHN_NO"


End Sub

Private Sub cmd9_Click(Index As Integer)
If (rk.BOF = True) Then
  rk.MoveLast
Else
  rk.MovePrevious

End Sub

Private Sub Command1_Click()
rk.Delete

MsgBox "Current Record Deleted"
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
cn.CursorLocation = adUseClient
cn.Open

Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open "department", cn, , , adCmdTable
MsgBox "loading please wait !!.............. ", , "Loading Message......."


Set Text1.DataSource = rk
Text1.DataField = "dept_id"
Set Text2.DataSource = rk
Text2.DataField = "DEPT_NAME"
Set Text4.DataSource = rk
Text4.DataField = "DEPT_DOCTOR"
Set Text5.DataSource = rk
Text5.DataField = "DEPT_HEAD"
Set Text6.DataSource = rk
Text6.DataField = "DEPT_SALARY"
Set Text7.DataSource = rk
Text7.DataField = "PHN_NO"


rk.MoveFirst



End Sub


Private Sub cmd2_Click(Index As Integer)
Unload departments
End Sub

