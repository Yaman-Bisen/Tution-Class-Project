VERSION 5.00
Begin VB.Form frm_update_student_finance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Student Information :"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10395
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   4620
      Left            =   7920
      TabIndex        =   13
      Top             =   840
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Caption         =   "Student Details :"
      Height          =   7815
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   7695
      Begin VB.CommandButton Command1 
         Caption         =   "&Update"
         Height          =   495
         Left            =   2160
         TabIndex        =   28
         Top             =   6600
         Width           =   1455
      End
      Begin VB.ComboBox Combo6 
         Height          =   405
         Left            =   4920
         TabIndex        =   22
         Text            =   "(Select Subject)"
         Top             =   1920
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   3240
         TabIndex        =   21
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   3240
         TabIndex        =   19
         Top             =   5520
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   3240
         TabIndex        =   17
         Top             =   4800
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   3240
         TabIndex        =   15
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   3240
         TabIndex        =   12
         Top             =   2640
         Width           =   1695
      End
      Begin VB.ComboBox Combo5 
         Height          =   405
         ItemData        =   "frm_update_student_finance.frx":0000
         Left            =   2400
         List            =   "frm_update_student_finance.frx":000D
         TabIndex        =   10
         Text            =   "(Select Package)"
         Top             =   1920
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
         Height          =   405
         Left            =   2400
         TabIndex        =   5
         Text            =   "(Select)"
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label12 
         Caption         =   "0"
         Height          =   375
         Left            =   5160
         TabIndex        =   25
         Top             =   5520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "0"
         Height          =   375
         Left            =   5160
         TabIndex        =   24
         Top             =   4800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "0"
         Height          =   375
         Left            =   5160
         TabIndex        =   23
         Top             =   4080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Balance Fees :"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Third Installment :"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   5520
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Second Installment :"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "First Installment :"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Total Fees :"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Package Type :"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Roll No :"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Select Student :"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.ComboBox Combo4 
         Height          =   405
         Left            =   4800
         TabIndex        =   6
         Text            =   "(Season)"
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frm_update_student_finance.frx":003E
         Left            =   2640
         List            =   "frm_update_student_finance.frx":005A
         TabIndex        =   2
         Text            =   "(Sem)"
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frm_update_student_finance.frx":00A0
         Left            =   120
         List            =   "frm_update_student_finance.frx":00B0
         TabIndex        =   1
         Text            =   "(Field)"
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   375
      Left            =   7920
      TabIndex        =   27
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   375
      Left            =   7920
      TabIndex        =   26
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frm_update_student_finance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim count1 As Integer
Dim count2 As Integer
Dim count3 As Integer
Dim count4 As Integer
Dim count5 As Integer

Private Sub Combo1_Click()
     If Combo1.Text = "OTHER" Then
        If count2 = 0 And Combo1.Text = "OTHER" Then
            count1 = 0
            Combo2.AddItem ("Sem-VII")
            Combo2.AddItem ("Sem-VIII")
        End If
    Else
        If count1 <> 1 Then
            count2 = 0
            count1 = 1
            Combo2.RemoveItem (7)
            Combo2.RemoveItem (6)
        End If
    End If

    Combo3.Clear
    If Combo2.Text <> "(Sem)" And Combo4.Text <> "(Season)" Then
        Set rs = cnn.Execute("select * from admission where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "' and season ='" & Combo4.Text & "'")
        If (Not rs.EOF) Then
            Do While Not rs.EOF
                Combo3.AddItem (rs.Fields("sname"))
                rs.MoveNext
            Loop
        Else
            MsgBox "Students not found", vbOKOnly + vbInformation, "Information"
        End If
    End If
End Sub

Private Sub Combo2_Click()
    If Combo4.Text <> "(Season)" Then
        Combo3.Clear
        If Combo1.Text = "(Field)" Then
            MsgBox "Please select the Field", vbQuestion + vbOKOnly, "Warning"
        ElseIf Combo2.Text = "(Sem)" Then
            MsgBox "Please select the Semester", vbQuestion + vbOKOnly, "Warning"
        Else
            Set rs = cnn.Execute("select * from admission where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "' and season ='" & Combo4.Text & "'")
            If (Not rs.EOF) Then
                Do While Not rs.EOF
                    Combo3.AddItem (rs.Fields("sname"))
                    rs.MoveNext
                Loop
            Else
                MsgBox "Students not found", vbOKOnly + vbInformation, "Information"
            End If
        End If
    End If
End Sub
Private Sub Combo3_Click()
    count3 = 0
    count4 = 0
    count5 = 0
    List1.Clear
    Set rs = cnn.Execute("Select * from admission where sname='" & Combo3.Text & "' and season='" & Combo4.Text & "'")
    If (Not rs.EOF) Then
        Label3.Caption = rs.Fields("rollno")
        Combo5.Text = rs.Fields("subject")
        Text1.Text = rs.Fields("totalfee")
        Set rs1 = cnn.Execute("select * from stud_sub_detail where sname='" & Combo3.Text & "'and season='" & Combo4.Text & "'and rollno='" & Label3.Caption & "' and sem='" & Combo2.Text & "'")
        If (Not rs1.EOF) Then
            If rs1.Fields("sub1") <> "Null" Then
                List1.AddItem (rs1.Fields("sub1"))
            End If
            If rs1.Fields("sub2") <> "Null" Then
                List1.AddItem (rs1.Fields("sub2"))
            End If
            If rs1.Fields("sub3") <> "Null" Then
                List1.AddItem (rs1.Fields("sub3"))
            End If
            If rs1.Fields("sub4") <> "Null" Then
                List1.AddItem (rs1.Fields("sub4"))
            End If
            If rs1.Fields("sub5") <> "Null" Then
                List1.AddItem (rs1.Fields("sub5"))
            End If
            If rs1.Fields("sub6") <> "Null" Then
                List1.AddItem (rs1.Fields("sub6"))
            End If
        End If
            Set rs2 = cnn.Execute("select * from trans where rollno='" & Label3.Caption & "'and sname='" & Combo3.Text & "'")
            If (Not rs2.EOF) Then
                If rs2.Fields("balance") = 0 Then
                    Text5.Text = "Paid"
                Else
                    Text5.Text = rs2.Fields("balance")
                End If
                If rs2.Fields("firstinstallment") <> 0 Then
                    Text2.Text = rs2.Fields("firstinstallment")
                End If
                If rs2.Fields("secondinstallment") <> 0 Then
                    Text3.Text = rs2.Fields("secondinstallment")
                    Label7.Visible = True
                    Text3.Visible = True
                Else
                    Text3.Text = 0
                End If
                If rs2.Fields("thirdinstallment") <> 0 Then
                    Text4.Text = rs2.Fields("thirdinstallment")
                    Label8.Visible = True
                    Text4.Visible = True
                Else
                    Text4.Text = 0
                End If
            End If
        If Combo5.Text <> "Select Subjects" Then
            Set rs = cnn.Execute("select * from package_details where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "' and package_type='" & Combo5.Text & "'")
            If (Not rs.EOF) Then
                Label10.Caption = rs.Fields("first_installment")
                Label11.Caption = rs.Fields("second_installment")
                Label12.Caption = rs.Fields("third_installment")
                If Label10.Caption = "0" Then
                    Label10.Visible = False
                Else
                    Label6.Visible = True
                    Text2.Visible = True
                    Label10.Visible = True
                End If
                If Label11.Caption = "0" Then
                    Label11.Visible = False
                Else
                    Label7.Visible = True
                    Text3.Visible = True
                    Label11.Visible = True
                End If
                If Label12.Caption = "0" Then
                    Label12.Visible = False
                Else
                    Label8.Visible = True
                    Text4.Visible = True
                    Label12.Visible = True
                End If
            End If
        Else
            Set rs = cnn.Execute("select * from trans where rollno='" & Label3.Caption & "'and sname='" & Combo3.Text & "'")
            If (Not rs.EOF) Then
                Label10.Visible = True
                Label11.Visible = True
                Label10.Caption = rs.Fields("oneinstallment")
                Label11.Caption = rs.Fields("twoinstallment")
                Label8.Visible = False
                Text4.Visible = False
                Label12.Visible = False
                If Label12.Caption = "0" Then
                    Label12.Visible = False
                Else
                    Label12.Visible = True
                End If
            End If
        End If
    End If
End Sub

Private Sub Combo4_Click()
    If Combo1.Text = "(Field)" Or Combo2.Text = "(Sem)" Or Combo4.Text = "(Season)" Then
        MsgBox "Please select all the information", vbExclamation + vbOKOnly, "Warning"
    Else
        Set rs = cnn.Execute("select * from admission where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "' and season ='" & Combo4.Text & "'")
        If (Not rs.EOF) Then
            Do While Not rs.EOF
                Combo3.AddItem (rs.Fields("sname"))
                rs.MoveNext
            Loop
        Else
            MsgBox "Students not found", vbOKOnly + vbInformation, "Information"
        End If
    End If
End Sub

Private Sub Combo5_Click()
    Combo6.Clear
    Combo6.Text = "(Select)"
    Set rs = cnn.Execute("select * from package_details where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "' and package_type='" & Combo5.Text & "'")
        If (Not rs.EOF) Then
            Label10.Caption = rs.Fields("first_installment")
            Label11.Caption = rs.Fields("second_installment")
            Label12.Caption = rs.Fields("third_installment")
            If Label12.Caption = "0" Then
                Label12.Visible = False
                Label8.Visible = False
                Text4.Visible = False
            Else
                Label8.Visible = True
                Text4.Visible = True
                Label12.Visible = True
            End If
        End If
    List1.Clear
    If Combo5.Text = "(Select Package)" Then
            MsgBox "Please select the package or subject", vbCritical + vbOKOnly, "Warning"
    ElseIf Combo5.Text = "All Package" Then
        List1.Visible = True
        Combo6.Visible = False
        Set rs = cnn.Execute("select * from package_details where field='" + Combo1.Text + "' and sem='" + Combo2.Text + "' and package_type='" + Combo5.Text + "'")
        Dim i As Integer
        i = 4
        If (Not rs.EOF) Then
                Text1.Text = rs.Fields("price")
                Text5.Text = Text1.Text
                While (i < 10)
                    If rs.Fields(i) <> "Null" Then
                        List1.AddItem (rs.Fields(i))
                    End If
                    i = i + 1
                Wend
        Else
            MsgBox "All Package is not available for selected semester", vbInformation + vbOKOnly, "Information"
        End If
        Text5.Text = Val(Text1.Text) - (Val(Text2.Text) - Val(Text3.Text) - Val(Text4.Text))
    ElseIf Combo5.Text = "Small Package" Then
        List1.Visible = True
        Combo6.Visible = False
        Set rs = cnn.Execute("select * from package_details where field='" + Combo1.Text + "' and sem='" + Combo2.Text + "' and package_type='" + Combo5.Text + "'")
        Dim j As Integer
        j = 4
        If (Not rs.EOF) Then
                Text1.Text = rs.Fields("price")
                Text5.Text = Text1.Text
                While (j < 10)
                    If rs.Fields(j) <> "Null" Then
                        List1.AddItem (rs.Fields(j))
                    End If
                    j = j + 1
                Wend
        Else
            MsgBox "Small Package is not available for selected semester", vbInformation + vbOKOnly, "Information"
        End If
        Text5.Text = Val(Text1.Text) - (Val(Text2.Text) - Val(Text3.Text) - Val(Text4.Text))
    Else
        Text5.Text = 0
        If Combo1.Text = "OTHER" Then
        Set rs = cnn.Execute("select distinct subject from subject_details order by subject")
        Do While Not rs.EOF
            Combo6.AddItem (rs.Fields("subject"))
            rs.MoveNext
        Loop
    Else
        Set rs = cnn.Execute("Select * from subject_details where years='" + Combo1.Text + "' and sem='" + Combo2.Text + "'")
        Do While Not rs.EOF
            Combo6.AddItem (rs.Fields("subject"))
            rs.MoveNext
        Loop
    End If
            Label10.Caption = ""
            Label11.Caption = ""
            Label12.Caption = ""
            Text1.Text = 0
            Combo6.Visible = True
            List1.Visible = True
            Label8.Visible = False
            Label12.Visible = False
            Text4.Visible = False
    End If
End Sub

Private Sub Combo6_Click()
If Combo5.Text <> "OTHER" Then
    If List1.ListCount = 2 Then
        MsgBox "Sorry two subjects are already selected", vbExclamation + vbOKOnly, "Information"
    ElseIf List1.ListCount = 0 Then
            List1.AddItem (Combo6.Text)
            Set rs = cnn.Execute("select * from subject_details where years='" + Combo1.Text + "' and sem='" + Combo2.Text + "'and subject='" + Combo6.Text + "'")
            If (Not rs.EOF) Then
                Dim a As Single
                a = rs.Fields("cost")
                Label13.Caption = a
                Text1.Text = Val(Text1.Text) + a
                Text5.Text = Val(Text5.Text) + a
            End If
    Else
        For m = 0 To List1.ListCount
            If List1.List(m) <> Combo6.Text Then
                If m = List1.ListCount Then
                    List1.AddItem (Combo6.Text)
                    Set rs = cnn.Execute("select * from subject_details where years='" + Combo1.Text + "' and sem='" + Combo2.Text + "'and subject='" + Combo6.Text + "'")
                    If (Not rs.EOF) Then
                        Dim a1 As Single
                        a1 = rs.Fields("cost")
                        Label14.Caption = a1
                        Text1.Text = Val(Text1.Text) + a1
                        Text5.Text = Val(Text5.Text) + a1
                    End If
                    Exit For
                End If
            Else
                MsgBox "Subject already selected", vbInformation + vbOKOnly, "Information"
                Exit For
            End If
        Next
    End If
End If
    If List1.ListCount = 1 Then
         Label10.Caption = Text1.Text
    ElseIf List1.ListCount = 2 Then
        If Label13.Caption > Label14.Caption Then
            Label10.Caption = Label13.Caption
            Label11.Caption = Label14.Caption
        Else
            Label10.Caption = Label14.Caption
            Label11.Caption = Label13.Caption
        End If
    End If
    Text5.Text = Val(Text1.Text) - (Val(Text2.Text) - Val(Text3.Text) - Val(Text4.Text))
End Sub

Private Sub Command1_Click()
    s = "update admission set subject='" & Combo5.Text & "',totalfee='" & Text1.Text & "' where rollno='" & Label3.Caption & "'and sname= '" & Combo3.Text & "'"
    cnn.Execute s
    
    If List1.ListCount = 1 Then
            m = "update stud_sub_detail set sub1='" & List1.List(0) & "',sub2='" & "Null" & "',sub3='" & "Null" & "',sub4='" & "Null" & "',sub5='" & "Null" & "',sub6='" & "Null" & "',package_type='" & Combo5.Text & "'where rollno='" & Label3.Caption & "'and sname= '" & Combo3.Text & "'"
            cnn.Execute m
    ElseIf List1.ListCount = 2 Then
            m = "update stud_sub_detail set sub1='" & List1.List(0) & "',sub2='" & List1.List(1) & "',sub3='" & "Null" & "',sub4='" & "Null" & "',sub5='" & "Null" & "',sub6='" & "Null" & "',package_type='" & Combo5.Text & "'where rollno='" & Label3.Caption & "'and sname= '" & Combo3.Text & "'"
            cnn.Execute m
    ElseIf List1.ListCount = 3 Then
            m = "update stud_sub_detail set sub1='" & List1.List(0) & "',sub2='" & List1.List(1) & "',sub3='" & List1.List(2) & "',sub4='" & "Null" & "',sub5='" & "Null" & "',sub6='" & "Null" & "',package_type='" & Combo5.Text & "'where rollno='" & Label3.Caption & "'and sname= '" & Combo3.Text & "'"
            cnn.Execute m
    ElseIf List1.ListCount = 4 Then
            m = "update stud_sub_detail set sub1='" & List1.List(0) & "',sub2='" & List1.List(1) & "',sub3='" & List1.List(2) & "',sub4='" & List1.List(3) & "',sub5='" & "Null" & "',sub6='" & "Null" & "',package_type='" & Combo5.Text & "'where rollno='" & Label3.Caption & "'and sname= '" & Combo3.Text & "'"
            cnn.Execute m
    ElseIf List1.ListCount = 5 Then
            m = "update stud_sub_detail set sub1='" & List1.List(0) & "',sub2='" & List1.List(1) & "',sub3='" & List1.List(2) & "',sub4='" & List1.List(3) & "',sub5='" & List1.List(4) & "',sub6='" & "Null" & "',package_type='" & Combo5.Text & "'where rollno='" & Label3.Caption & "'and sname= '" & Combo3.Text & "'"
            cnn.Execute m
    ElseIf List1.ListCount = 6 Then
            m = "update stud_sub_detail set sub1='" & List1.List(0) & "',sub2='" & List1.List(1) & "',sub3='" & List1.List(2) & "',sub4='" & List1.List(3) & "',sub5='" & List1.List(4) & "',sub6='" & List1.List(5) & "',package_type='" & Combo5.Text & "'where rollno='" & Label3.Caption & "'and sname= '" & Combo3.Text & "'"
            cnn.Execute m
    End If
    If Combo5.Text = "Select Subjects" Then
        n = "update trans set oneinstallment='" & Label10.Caption & "',twoinstallment='" & Label11.Caption & "',firstinstallment='" & Text2.Text & "',secondinstallment='" & Text3.Text & "',balance='" & Text5.Text & "'where rollno='" & Label3.Caption & "'and sname= '" & Combo3.Text & "'"
        cnn.Execute n
        MsgBox "Student information updated successfully", vbInformation + vbOKOnly, "Information"
        Unload Me
    Else
        n = "update trans set firstinstallment='" & Text2.Text & "',secondinstallment='" & Text3.Text & "',thirdinstallment='" & Text4.Text & "',balance='" & Text5.Text & "'where rollno='" & Label3.Caption & "'and sname= '" & Combo3.Text & "'"
        cnn.Execute n
        MsgBox "Student information updated successfully", vbInformation + vbOKOnly, "Information"
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 700
    Me.Left = 1500
    connect
    count1 = 0
    count2 = 1
    count3 = 0
    a = Date
    b = Mid(a, 9, 10)
    Combo4.AddItem ("Winter " + b)
    Combo4.AddItem ("Summer " + b)
End Sub



Private Sub Text2_Change()
    If count3 = 0 Then
        count3 = 1
    Else
        If Val(Text2.Text) <= Val(Label10.Caption) Then
            Text5.Text = Val(Text1.Text) - Val(Text2.Text)
        Else
            MsgBox "Value can't be greater than first installment", vbCritical + vbOKOnly, "Warning"
            Text2.Text = Label10.Caption
        End If
    End If
End Sub

Private Sub Text3_Change()
    If count4 = 0 Then
        count4 = 1
    Else
        If Val(Text2.Text) = Val(Label10.Caption) Then
            If Val(Text3.Text) <= Val(Label11.Caption) Then
                Text5.Text = Val(Text1.Text) - Val(Text3.Text)
            Else
                MsgBox "Value can't be greater than second installment", vbCritical + vbOKOnly, "Warning"
                Text3.Text = Label11.Caption
            End If
        Else
            MsgBox "Please complete first installment first", vbExclamation + vbOKOnly, "Warning"
            Text2.SetFocus
            Text3.Text = 0
        End If
    End If
End Sub

Private Sub Text4_Change()
     If count5 = 0 Then
        count5 = 1
    Else
        If Val(Text2.Text) = Val(Label10.Caption) And Val(Text3.Text) = Val(Label11.Caption) Then
            If Val(Text4.Text) <= Val(Label12.Caption) Then
                Text5.Text = Val(Text1.Text) - Val(Text4.Text)
            Else
                MsgBox "Value can't be greater than third installment", vbCritical + vbOKOnly, "Warning"
                Text4.Text = Label12.Caption
            End If
        Else
            MsgBox "Please complete remaining installments first", vbExclamation + vbOKOnly, "Warning"
            Text4.Text = 0
        End If
    End If
End Sub
