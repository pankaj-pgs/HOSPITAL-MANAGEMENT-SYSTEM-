VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form addmedicines 
   Caption         =   "Medicine Store"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   15975
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Main"
      Height          =   10815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   20175
      Begin VB.Frame Frame3 
         BackColor       =   &H0000C0C0&
         Caption         =   "Control Buttons"
         Height          =   1335
         Left            =   600
         TabIndex        =   9
         Top             =   8760
         Width           =   19095
         Begin VB.CommandButton Command2 
            Height          =   495
            Left            =   1800
            Picture         =   "addmedicines.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Left            =   3240
            Picture         =   "addmedicines.frx":58AB
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd8 
            Height          =   495
            Index           =   8
            Left            =   11640
            Picture         =   "addmedicines.frx":B03A
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd10 
            Height          =   495
            Index           =   6
            Left            =   13080
            Picture         =   "addmedicines.frx":108B6
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd6 
            Height          =   495
            Index           =   5
            Left            =   10200
            Picture         =   "addmedicines.frx":1617D
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd3 
            Height          =   495
            Index           =   4
            Left            =   6000
            Picture         =   "addmedicines.frx":1BA5A
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd4 
            Height          =   495
            Index           =   3
            Left            =   7320
            Picture         =   "addmedicines.frx":212E4
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd5 
            Height          =   495
            Index           =   2
            Left            =   8760
            Picture         =   "addmedicines.frx":26B8F
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd9 
            Height          =   495
            Index           =   1
            Left            =   4560
            Picture         =   "addmedicines.frx":2C419
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd2 
            Height          =   495
            Index           =   0
            Left            =   14520
            Picture         =   "addmedicines.frx":31CA0
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0000C0C0&
         Caption         =   "Medicine  Information"
         Height          =   7695
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   19095
         Begin SHDocVwCtl.WebBrowser WebBrowser1 
            Height          =   2055
            Left            =   11400
            TabIndex        =   29
            Top             =   360
            Width           =   4575
            ExtentX         =   8070
            ExtentY         =   3625
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
            Location        =   ""
         End
         Begin VB.PictureBox Picture2 
            Height          =   2415
            Left            =   5400
            Picture         =   "addmedicines.frx":37557
            ScaleHeight     =   2355
            ScaleWidth      =   5475
            TabIndex        =   28
            Top             =   4800
            Width           =   5535
         End
         Begin VB.PictureBox Picture1 
            Height          =   4815
            Left            =   11400
            Picture         =   "addmedicines.frx":3CBD8
            ScaleHeight     =   4755
            ScaleWidth      =   7395
            TabIndex        =   27
            Top             =   2760
            Width           =   7455
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   2280
            TabIndex        =   24
            Top             =   3960
            Width           =   3975
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   2280
            TabIndex        =   23
            Top             =   3360
            Width           =   3975
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   2280
            TabIndex        =   22
            Top             =   2760
            Width           =   3975
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   2280
            TabIndex        =   21
            Top             =   2160
            Width           =   3975
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   2280
            TabIndex        =   20
            Top             =   1560
            Width           =   3975
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   2280
            TabIndex        =   19
            Top             =   960
            Width           =   3975
         End
         Begin VB.TextBox Text1 
            Height          =   405
            Left            =   2280
            TabIndex        =   18
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Medicine  Id"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Medicine Name"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Manufacturer  Name"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   6
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Type of Medicine"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   5
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Batch Number"
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   4
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Manufacturee Date"
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   3
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Date Of Expiry"
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   2
            Top             =   3960
            Width           =   1575
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   12240
      Picture         =   "addmedicines.frx":42F6C
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1920
   End
End
Attribute VB_Name = "addmedicines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd10_Click(Index As Integer)
Dim medicineno As Integer
medicineno = 1
sqk = "SELECT med_code FROM (SELECT med_code FROM medicine ORDER BY med_code desc) WHERE ROWNUM <= 1 "

Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open sqk, cn, , , adCmdText
medicineno = rk.Fields(0)
Text1.Text = medicineno + 1

         
 rk.Close
         
         Text2.Text = " "
         Text3.Text = " "
         Text7.Text = " "
         Text6.Text = " "
         Text5.Text = " "
         Text4.Text = " "
End Sub



Private Sub cmd2_Click(Index As Integer)
Unload addmedicines
End Sub

Private Sub cmd3_Click(Index As Integer)
On Error GoTo errordesc
rk.MoveFirst
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call medicine
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
Call medicine
End Sub

Private Sub cmd5_Click(Index As Integer)
On Error GoTo errordesc
rk.MoveLast
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call medicine
End Sub

Private Sub cmd6_Click(Index As Integer)
On Error GoTo errordesc
If (rk.BOF = True) Then
    rk.MoveLast
Else
    rk.MovePrevious
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call medicine
End Sub

Private Sub cmd8_Click(Index As Integer)
Dim aid As String
aid = InputBox("Enter the doctor id to search", "search", "id")
sqk = "select * from doctor where doctor_id = " & Val(aid)
Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open sqk, cn, , , adCmdText
If rk.EOF = True Then
MsgBox "record not found!!..................."
Call medicine
End If
Set Text2.DataSource = rk
Text2.DataField = "MED_NAME"
Set Text4.DataSource = rk
Text4.DataField = "MED_TYPE"
Set Text3.DataSource = rk
Text3.DataField = "MANUFACT_NAME"
Set Text6.DataSource = rk
Text6.DataField = "MFG_DATE"
Set Text7.DataSource = rk
Text7.DataField = "EXP_DATE"

Text1.DataField = "MED_CODE"
Set Text4.DataSource = rk
Text4.DataField = "batch_no"
End Sub

Private Sub cmd9_Click(Index As Integer)
On Error GoTo errordesc
rk.AddNew
rk.Update
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call medicine
End Sub
Private Sub medicine()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Open "medicine", cn, , , adCmdTable

Set rk = rs
Set Text2.DataSource = rk
Text2.DataField = "MED_NAME"
Set Text4.DataSource = rk
Text4.DataField = "MED_TYPE"
Set Text3.DataSource = rk
Text3.DataField = "MANUFACT_NAME"
Set Text6.DataSource = rk
Text6.DataField = "MFG_DATE"
Set Text7.DataSource = rk
Text7.DataField = "EXP_DATE"
Set Text1.DataSource = rk
Text1.DataField = "MED_CODE"
Set Text5.DataSource = rk
Text5.DataField = "batch_no"
End Sub

Private Sub Command1_Click()
rk.Delete
End Sub

Private Sub Command2_Click()
rk.Update
MsgBox " UPDATION OF DATA COMPLETED"
rk.MoveFirst
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate "C:\hospital project\images\tablettak.gif"

Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
cn.CursorLocation = adUseClient
cn.Open

Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open "medicine", cn, , , adCmdTable
MsgBox "loading please wait !!.............. ", , "Loading Message......."


Set Text2.DataSource = rk
Text2.DataField = "MED_NAME"
Set Text4.DataSource = rk
Text4.DataField = "MED_TYPE"
Set Text3.DataSource = rk
Text3.DataField = "MANUFACT_NAME"
Set Text6.DataSource = rk
Text6.DataField = "MFG_DATE"
Set Text7.DataSource = rk
Text7.DataField = "EXP_DATE"
Set Text1.DataSource = rk
Text1.DataField = "MED_CODE"
Set Text5.DataSource = rk
Text5.DataField = "batch_no"

End Sub

