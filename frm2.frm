VERSION 5.00
Begin VB.MDIForm frmmain 
   BackColor       =   &H8000000C&
   Caption         =   " HOSPITAL MANAGEMENT"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11400
   LinkTopic       =   "MDIForm1"
   Picture         =   "frm2.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu patient 
      Caption         =   "Patient"
      Begin VB.Menu addpatient 
         Caption         =   "AddPatient"
      End
      Begin VB.Menu updatepatient 
         Caption         =   "UpdatePatient"
      End
      Begin VB.Menu dischargepatient 
         Caption         =   "DischargePatient"
      End
   End
   Begin VB.Menu doctor 
      Caption         =   "Doctor"
      Begin VB.Menu adddoctor 
         Caption         =   "AddDoctor"
      End
      Begin VB.Menu updatedoctor 
         Caption         =   "UpdateDoctor"
      End
      Begin VB.Menu deletedoctor 
         Caption         =   "DeleteDoctor"
      End
   End
   Begin VB.Menu nurse 
      Caption         =   "Nurse"
      Begin VB.Menu AddNurse 
         Caption         =   "AddNurse"
      End
      Begin VB.Menu UpdateNurse 
         Caption         =   "UpdateNurse"
      End
      Begin VB.Menu RemoveNurse 
         Caption         =   "RemoveNurse"
      End
   End
   Begin VB.Menu worktim 
      Caption         =   "WorkTime"
      Begin VB.Menu addworktime 
         Caption         =   "Employee Worktime"
      End
   End
   Begin VB.Menu department 
      Caption         =   "Department"
      Begin VB.Menu departmentdetail 
         Caption         =   "DepartmentDetail"
      End
   End
   Begin VB.Menu medicines 
      Caption         =   "Medicine"
      Begin VB.Menu addmedicine 
         Caption         =   "AddMedicine"
      End
      Begin VB.Menu updatemedicine 
         Caption         =   "UpdateMedicine"
      End
      Begin VB.Menu deletemedicine 
         Caption         =   "DeleteMedicine"
      End
   End
   Begin VB.Menu report 
      Caption         =   "Report"
      Begin VB.Menu doctors 
         Caption         =   "Doctors"
      End
      Begin VB.Menu patients 
         Caption         =   "Patients"
      End
      Begin VB.Menu medicine 
         Caption         =   "Medicine"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cn As ADODB.Connection
Public cmd1 As String
Public rs As ADODB.Recordset
Public rk As ADODB.Recordset
Public i, j, variable As Integer
 
Public sqlcmd, sqk, strsql, strname, str1, rate, a, b, c As String

Private Sub about_Click()
frmAbout.Show

End Sub

Private Sub addattendance_Click()
employeeattendance.Show

End Sub

Private Sub adddoctor_Click()
adddoctors.Show

End Sub

Private Sub addmedicine_Click()
addmedicines.Show

End Sub

Private Sub AddNurse_Click()
addnurses.Show
End Sub

Private Sub addpatient_Click()
addmissionpatient.Show

End Sub

Private Sub addworktime_Click()
employeeattendance.Show

End Sub

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub deleteattendance_Click()
employeeattendance
End Sub

Private Sub deletedoctor_Click()
removedoctor.Show
End Sub

Private Sub deletemedicine_Click()
addmedicines.Show
End Sub

Private Sub deleteworktime_Click()
employeeworktime.Show
End Sub

Private Sub departmentdetail_Click()
departments.Show

End Sub

Private Sub detail_Click()
reportemployee.Show

End Sub

Private Sub dischargepatient_Click()
dischargpatient.Show
End Sub

Private Sub doctors_Click()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
cn.CursorLocation = adUseClient
cn.Open
Dim aid As String
aid = InputBox("Enter the doctor id to search", "search", "id")
Set rk = New ADODB.Recordset
        rk.CursorType = adOpenDynamic
        rk.LockType = adLockOptimistic
        sqk = " create or replace view  reportdoctor as select e.DOCT_NAME,e. D_O_B,e. SALARY,e.PHN_NO,e.EMAIL_ID,e.DOCTOR_ID,e. SEX,e.DEPT_ID from  doctor e where e.doctor_ID = " & Val(aid)
        rk.Open sqk, cn, 2, 3

doctorReport.Show

End Sub

Private Sub exit_Click()
If MsgBox("Do you really want to exit? press yes or no", vbYesNo, "Hospital Management System") = vbYes Then
Unload frmmain

End If

End Sub

Private Sub medicine_Click()

reportmedicine.Show

End Sub

Private Sub patients_Click()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
cn.CursorLocation = adUseClient
cn.Open
Dim aid As String
aid = InputBox("Enter the doctor id to search", "search", "id")
Set rk = New ADODB.Recordset
        rk.CursorType = adOpenDynamic
        rk.LockType = adLockOptimistic
        sqk = " create or replace view report as select e.PATIENT_ID,e.PAT_NAME,e.DOC_FEE,e.BLD_GRP,e.DISEASE,e. SEX,r.doctor_id,r.app_date,r.PRESCRIPTION,d.doct_name,d.QUALIFICATION  from patient e,appointment r,doctor d where d.doctor_id=r.doctor_id and e.PATIENT_ID ='" & Val(aid) & "'  and r.PATient_ID=" & Val(aid)
        rk.Open sqk, cn, 2, 3
PatientReport.Show
End Sub

Private Sub RemoveNurse_Click()
addnurses.Show
End Sub

Private Sub updateattendance_Click()
employeeattendance
End Sub

Private Sub updatedoctor_Click()
removedoctor.Show

End Sub

Private Sub updatemedicine_Click()
addmedicines.Show
End Sub

Private Sub UpdateNurse_Click()
addnurses.Show
End Sub

Private Sub updatepatient_Click()
addmissionpatient.Show
End Sub

Private Sub updateworktime_Click()
employeeworktime.Show
End Sub
