VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form adddoctors 
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Main"
      Height          =   10815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   20055
      Begin VB.Frame Frame3 
         BackColor       =   &H0000C0C0&
         Caption         =   "Control Buttons"
         Height          =   1335
         Left            =   720
         TabIndex        =   11
         Top             =   8880
         Width           =   18735
         Begin VB.CommandButton Command2 
            Height          =   495
            Left            =   1800
            Picture         =   "adddoctor.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   3360
            Picture         =   "adddoctor.frx":58C7
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmd8 
            Height          =   495
            Index           =   8
            Left            =   13080
            Picture         =   "adddoctor.frx":B172
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmd6 
            Height          =   495
            Index           =   5
            Left            =   11520
            Picture         =   "adddoctor.frx":109EE
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmd3 
            Height          =   495
            Index           =   4
            Left            =   6600
            Picture         =   "adddoctor.frx":162CB
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmd4 
            Height          =   495
            Index           =   3
            Left            =   8160
            Picture         =   "adddoctor.frx":1BB55
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmd5 
            Height          =   495
            Index           =   2
            Left            =   9840
            Picture         =   "adddoctor.frx":21400
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmd9 
            Height          =   495
            Index           =   1
            Left            =   4920
            Picture         =   "adddoctor.frx":26C8A
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmd2 
            Height          =   495
            Index           =   0
            Left            =   120
            Picture         =   "adddoctor.frx":2C511
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0000C0C0&
         Caption         =   "Doctor information"
         Height          =   7695
         Left            =   720
         TabIndex        =   1
         Top             =   600
         Width           =   18735
         Begin VB.PictureBox Picture1 
            Height          =   6015
            Left            =   7080
            Picture         =   "adddoctor.frx":31DC8
            ScaleHeight     =   5955
            ScaleWidth      =   4515
            TabIndex        =   33
            Top             =   240
            Width           =   4575
         End
         Begin SHDocVwCtl.WebBrowser WebBrowser1 
            Height          =   975
            Left            =   13440
            TabIndex        =   31
            Top             =   360
            Width           =   3375
            ExtentX         =   5953
            ExtentY         =   1720
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
         Begin VB.TextBox Text9 
            Height          =   375
            Left            =   2280
            TabIndex        =   29
            Top             =   6000
            Width           =   3855
         End
         Begin VB.TextBox Text8 
            Height          =   375
            Left            =   2280
            TabIndex        =   28
            Top             =   5400
            Width           =   3855
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   2280
            TabIndex        =   27
            Top             =   4800
            Width           =   3855
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   2280
            TabIndex        =   26
            Top             =   4200
            Width           =   3855
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   2280
            TabIndex        =   25
            Top             =   3600
            Width           =   3855
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   2280
            TabIndex        =   24
            Top             =   3000
            Width           =   3855
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   2280
            TabIndex        =   23
            Top             =   2280
            Width           =   3855
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   2280
            TabIndex        =   22
            Top             =   840
            Width           =   3855
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   2280
            TabIndex        =   21
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "adddoctor.frx":616DC
            Left            =   2280
            List            =   "adddoctor.frx":616E6
            TabIndex        =   19
            Text            =   "Male"
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail id"
            Height          =   375
            Index           =   9
            Left            =   360
            TabIndex        =   20
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Gender"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   10
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Id"
            BeginProperty Font 
               Name            =   "Berlin Sans FB Demi"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Name"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Eduqualification"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   7
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Date of Birth"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   6
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Department"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   5
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Patient Salary"
            Height          =   375
            Index           =   5
            Left            =   360
            TabIndex        =   4
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Fathers Name"
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   3
            Top             =   4920
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Phone Number"
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   2
            Top             =   5520
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "adddoctors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub doctor()


Set rs = New ADODB.Recordset
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Open "doctor", cn, , , adCmdTable
Set rk = rs

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
Combo1.DataField = "SEX"
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
Call doctor
End Sub

Private Sub cmd4_Click(Index As Integer)
On Error GoTo errordesc
If (rk.EOF = True) Then
    rk.MoveFirst
Else
    rk.MoveNext

End If
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call doctor

End Sub

Private Sub cmd5_Click(Index As Integer)
 On Error GoTo errordesc
 rk.MoveLast
 Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call doctor
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
Call doctor
End Sub

Private Sub cmd8_Click(Index As Integer)
On Error GoTo errordesc
Dim aid As String
aid = InputBox("Enter the doctor id to search", "search", "id")
sqk = "select * from doctor where doctor_id = " & Val(aid)

Set rk = New ADODB.Recordset
Set rs = New ADODB.Recordset

rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open sqk, cn, , , adCmdText

rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Open "doctor", cn, , , adCmdTable


If rk.EOF = True Then
MsgBox " doctor id not exist"
Set rk = rs

End If

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
Combo1.DataField = "SEX"
Exit Sub

errordesc:
i = MsgBox(Err.Description, vbCritical)
Call doctor
End Sub

Private Sub cmd9_Click(Index As Integer)
        Set rk = New ADODB.Recordset
        rk.CursorType = adOpenDynamic
        rk.LockType = adLockOptimistic
        sqk = "select * from doctor "
        rk.Open sqk, cn, , , adCmdText
        rk.AddNew
        rk.Fields(0) = Text2.Text
        rk.Fields(1) = Text4.Text
        rk.Fields(2) = Text3.Text
        rk.Fields(3) = Text6.Text
        rk.Fields(4) = Text7.Text
        rk.Fields(5) = Text8.Text
        rk.Fields(6) = Text9.Text
        rk.Fields(7) = Text1.Text
        rk.Fields(8) = Text5.Text
        rk.Fields(9) = Combo1.Text

        rk.Update
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
        Combo1.DataField = "SEX"

        
        rk.MoveFirst
        
        
End Sub

Private Sub Command1_Click()
Dim aid As String
aid = InputBox("Enter the doctor id to search", "search", "id")
sqk = "select * from doctor where doctor_id = " & Val(aid)
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
Combo1.DataField = "SEX"
rk.Update
rk.MoveFirst
End Sub

Private Sub Command2_Click()
 Dim doctorno As Integer
 doctorno = 1
sqk = "SELECT doctor_id FROM (SELECT doctor_id FROM doctor ORDER BY doctor_id desc) WHERE ROWNUM <= 1 "

Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open sqk, cn, , , adCmdText
doctorno = rk.Fields(0)
Text1.Text = doctorno + 1

         
 rk.Close
Text2.Text = ""
Text4.Text = ""
 Text3.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text5.Text = ""
Combo1.Text = ""
End Sub

Private Sub Form_Load()
On Error Resume Next
WebBrowser1.Navigate "C:\file.gif"
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
Combo1.DataField = "SEX"




End Sub


Private Sub Text1_LostFocus()
Dim aid As Integer
Dim z As Integer
aid = Text1.Text
sqk = "select * from doctor where doctor_id = " & Val(aid)
Set rs = New ADODB.Recordset
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Open sqk, cn, , , adCmdText
If (rs.EOF = True) Then
Else
z = rs.Fields(7)
End If
    If (z = aid) Then
        MsgBox "Enter Different Id, this id already exist!!........"
        Text1.Text = ""
        Text1.SetFocus
    End If
End Sub

