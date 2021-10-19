VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "D'Soft Computer Training Center"
   ClientHeight    =   6825
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11730
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_admission 
      Caption         =   "&Admission"
      Begin VB.Menu mnu_new_stud 
         Caption         =   "New Student"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_old_stud 
         Caption         =   "Old Student"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_seat_no 
         Caption         =   "Allocate Seat No"
      End
      Begin VB.Menu mnu_update_stud_inform 
         Caption         =   "Update Student Personal Inform..."
      End
      Begin VB.Menu mnu_update_install 
         Caption         =   "Update Student Installment Details ..."
      End
   End
   Begin VB.Menu mnu_installments 
      Caption         =   "&Pay Installments"
      Begin VB.Menu mnu_pay_installments 
         Caption         =   "Installments"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnu_add_college 
         Caption         =   "Add College"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnu_view 
      Caption         =   "&Reports"
      Begin VB.Menu mnu_college_stud_in_batch 
         Caption         =   "College Student In Batch"
      End
      Begin VB.Menu mnu_stud_batch 
         Caption         =   "Count Students in Batch"
      End
      Begin VB.Menu mnu_admitted 
         Caption         =   "All Admitted Students"
      End
      Begin VB.Menu mnu_batch_wise 
         Caption         =   "All Student In Batch"
      End
      Begin VB.Menu mnu_college 
         Caption         =   "All Student In College"
      End
      Begin VB.Menu mnu_1st_installment 
         Caption         =   "Installment pending"
      End
      Begin VB.Menu mnu_view_packages 
         Caption         =   "View packages"
      End
      Begin VB.Menu mnu_view_batches 
         Caption         =   "View Batches"
      End
   End
   Begin VB.Menu mnu_batch 
      Caption         =   "&Batch"
      Begin VB.Menu mnu_create_batch 
         Caption         =   "Create Batch"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnu_add_stud_batch 
         Caption         =   "Add Students to Batch"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnu_late 
         Caption         =   "Late Student to add"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnu_remove 
         Caption         =   "Remove Student From Batch"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnu_search 
      Caption         =   "&Search"
      Begin VB.Menu mnu_student_search 
         Caption         =   "Search Student"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnu_add 
      Caption         =   "A&dd"
      Begin VB.Menu mnu_add_subject 
         Caption         =   "Add Subject"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu_add_package 
         Caption         =   "Add Package"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnu_update 
      Caption         =   "&Update"
      Begin VB.Menu mnu_subject_price 
         Caption         =   "Update Subject Price"
      End
      Begin VB.Menu update_package_details 
         Caption         =   "Update Package Details"
      End
   End
   Begin VB.Menu mnu_finance 
      Caption         =   "&Finance"
      Begin VB.Menu mnu_date_report 
         Caption         =   "Date wise Report"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu_finance_rpt 
         Caption         =   "Whole Finance Report"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnu_switch 
      Caption         =   "S&witch"
      Begin VB.Menu mnu_admin 
         Caption         =   "To Admin"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu_user 
         Caption         =   "To User"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnu_quit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
   
    If login_type = "admin" Then
        mnu_add.Enabled = True
        mnu_update.Enabled = True
        mnu_admin.Visible = False
        mnu_user.Visible = True
        mnu_finance.Enabled = True
    ElseIf login_type = "user" Then
         mnu_add.Enabled = False
        mnu_update.Enabled = False
         mnu_admin.Visible = True
        mnu_user.Visible = False
        mnu_finance.Enabled = False
    End If
End Sub

Private Sub mnu_1st_installment_Click()
    frm_installment_bal_rpt.Show
End Sub


Private Sub mnu_add_college_Click()
    frm_add_college.Show
End Sub

Private Sub mnu_add_package_Click()
    frm_insert_package.Show
End Sub

Private Sub mnu_add_stud_batch_Click()
    frm_student_batch.Show
End Sub

Private Sub mnu_add_subject_Click()
    frm_insert_subject.Show
End Sub


Private Sub mnu_admin_Click()
    login_type = "admin"
    frm_login.Show
    Unload Me
End Sub

Private Sub mnu_admitted_Click()
    frm_admitted_stud_report.Show
End Sub


Private Sub mnu_batch_wise_Click()
    frm_batch_wise.Show
End Sub

Private Sub mnu_college_Click()
    frm_college_rpt.Show
End Sub

Private Sub mnu_college_stud_in_batch_Click()
    frm_batch_student.Show
End Sub

Private Sub mnu_create_batch_Click()
    frm_new_batch.Show
End Sub

Private Sub mnu_date_report_Click()
    frm_date_wise_finance.Show
End Sub

Private Sub mnu_finance_rpt_Click()
    Form2.Show
End Sub

Private Sub mnu_late_Click()
    frm_late_come_batch.Show
End Sub

Private Sub mnu_new_stud_Click()
    frm_new_admit_student.Show
End Sub

Private Sub mnu_old_stud_Click()
    Form1.Show
End Sub

Private Sub mnu_pay_installments_Click()
    frm_pay_installment.Show
End Sub


Private Sub mnu_quit_Click()
    End
End Sub

Private Sub mnu_remove_Click()
    frm_remove_stud_fro_batch.Show
End Sub

Private Sub mnu_seat_no_Click()
    frm_upadate_seat_no.Show
End Sub

Private Sub mnu_stud_batch_Click()
    frm_count_student_rpt.Show
End Sub

Private Sub mnu_student_search_Click()
    frm_search_student.Show
End Sub

Private Sub mnu_subject_price_Click()
    frm_update_sub_price.Show
End Sub

Private Sub mnu_update_install_Click()
    frm_update_student_finance.Show
End Sub

Private Sub mnu_update_stud_inform_Click()
    frm_update_student_inform.Show
End Sub

Private Sub mnu_user_Click()
    login_type = "user"
    frm_login.Show
    Unload Me
End Sub

Private Sub mnu_view_batches_Click()
    frm_batch_view.Show
End Sub

Private Sub mnu_view_packages_Click()
    frm_package_details.Show
End Sub

Private Sub mnu_whole_Click()
    Form2.Show
End Sub

Private Sub update_package_details_Click()
    frm_update_package_details.Show
End Sub
