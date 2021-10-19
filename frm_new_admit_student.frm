VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_new_admit_student 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Student Admission :"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12870
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10560
      TabIndex        =   32
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Details"
      Height          =   495
      Left            =   6360
      TabIndex        =   14
      Top             =   7560
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   4050
      Left            =   10200
      TabIndex        =   28
      Top             =   2160
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Student Detail :"
      Height          =   8175
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10215
      Begin VB.Frame Frame3 
         Height          =   1695
         Left            =   6000
         TabIndex        =   39
         Top             =   4680
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox Text9 
            Height          =   405
            Left            =   1800
            TabIndex        =   9
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Add"
            Height          =   375
            Left            =   1200
            TabIndex        =   40
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label26 
            Caption         =   "Enter Price :"
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.ComboBox Combo6 
         Height          =   405
         Left            =   7440
         TabIndex        =   2
         Text            =   "(Select)"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   8
         Top             =   4800
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   3000
         TabIndex        =   7
         Top             =   4200
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   3000
         TabIndex        =   6
         Top             =   3720
         Width           =   3255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Proceed"
         Enabled         =   0   'False
         Height          =   495
         Left            =   8280
         TabIndex        =   15
         Top             =   7560
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   6840
         TabIndex        =   1
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   159514625
         CurrentDate     =   43854
      End
      Begin VB.ComboBox Combo5 
         Enabled         =   0   'False
         Height          =   405
         Left            =   2880
         TabIndex        =   13
         Text            =   "(Select)"
         Top             =   7200
         Width           =   2655
      End
      Begin VB.ComboBox Combo4 
         Enabled         =   0   'False
         Height          =   405
         ItemData        =   "frm_new_admit_student.frx":0000
         Left            =   2880
         List            =   "frm_new_admit_student.frx":000D
         TabIndex        =   12
         Text            =   "(Select)"
         Top             =   6600
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frm_new_admit_student.frx":003E
         Left            =   2400
         List            =   "frm_new_admit_student.frx":004E
         TabIndex        =   0
         Text            =   "(Select)"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   405
         Left            =   3000
         TabIndex        =   3
         Top             =   1920
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   405
         ItemData        =   "frm_new_admit_student.frx":006E
         Left            =   2880
         List            =   "frm_new_admit_student.frx":0070
         TabIndex        =   10
         Text            =   "(Select College)"
         Top             =   5400
         Width           =   2655
      End
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         Height          =   405
         ItemData        =   "frm_new_admit_student.frx":0072
         Left            =   2880
         List            =   "frm_new_admit_student.frx":008E
         TabIndex        =   11
         Text            =   "(Select Sem)"
         Top             =   6000
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   405
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   5
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   405
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label20 
         Caption         =   "Select Season :"
         Height          =   375
         Left            =   5400
         TabIndex        =   38
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "Pincode :"
         Height          =   375
         Left            =   360
         TabIndex        =   37
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label Label18 
         Caption         =   "Land-Mark :"
         Height          =   375
         Left            =   360
         TabIndex        =   36
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Label Label17 
         Caption         =   "Area :"
         Height          =   375
         Left            =   360
         TabIndex        =   35
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Date :"
         Height          =   255
         Left            =   5760
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Select Subject :"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   7320
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Select Pakcage :"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   6720
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Select Field :"
         Height          =   375
         Left            =   480
         TabIndex        =   26
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Roll No :"
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         Height          =   375
         Left            =   2160
         TabIndex        =   24
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Student Name :"
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Personal Mobile No :"
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Parents Mobile No :"
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Select College :"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   5520
         Width           =   2535
      End
      Begin VB.Label Label10 
         Caption         =   "Select Semester :"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   6120
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Sr No :"
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Label Label16 
      Caption         =   "0"
      Height          =   375
      Left            =   10680
      TabIndex        =   34
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "0"
      Height          =   375
      Left            =   10680
      TabIndex        =   33
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      Height          =   255
      Left            =   10680
      TabIndex        =   29
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frm_new_admit_student"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim count1 As Integer
Dim count2 As Integer

Private Sub Combo1_Click()
    Dim a As String
    Dim b As String
    Dim c As String
    Dim d As String
    If Combo1.Text = "(Select)" Then
        MsgBox "Please select the field", vbCritical + vbOKOnly, "warning"
    Else
            Set rs = cnn.Execute("select * from admission where field='" + Combo1.Text + "'")
            Do While Not rs.EOF
                a = rs.Fields("rollno")
                rs.MoveNext
            Loop
            If Combo1.Text = "BSC-IT" Then
                b = Mid(a, 4)
                c = Val(b) + 1
                d = Mid(a, 1, 3)
                Label5.Caption = d + c
            Else
                b = Mid(a, 3)
                c = Val(b) + 1
                d = Mid(a, 1, 2)
                Label5.Caption = d + c
            End If
    End If
    If (Not Combo1.Text = "(Select)") Then
        Text1.Enabled = True
    End If
    If Combo1.Text = "OTHER" Then
        Frame3.Visible = True
    Else
        Frame3.Visible = False
    End If
    
    If Combo1.Text = "OTHER" Then
        If count2 = 0 And Combo1.Text = "OTHER" Then
            count1 = 0
            Combo3.AddItem ("Sem-VII")
            Combo3.AddItem ("Sem-VIII")
        End If
    Else
        If count1 <> 1 Then
            count2 = 0
            count1 = 1
            Combo3.RemoveItem (7)
            Combo3.RemoveItem (6)
        End If
    End If
    
End Sub

Private Sub Combo3_Click()
    Combo5.Clear
    List1.Clear
    Label12.Caption = 0
    Label16.Caption = 0
    Label15.Caption = 0
    If Combo1.Text = "OTHER" Then
        Set rs = cnn.Execute("select distinct subject from subject_details order by subject")
        Do While Not rs.EOF
            Combo5.AddItem (rs.Fields("subject"))
            rs.MoveNext
        Loop
    Else
        Set rs = cnn.Execute("Select * from subject_details where years='" + Combo1.Text + "' and sem='" + Combo3.Text + "'")
        Do While Not rs.EOF
            Combo5.AddItem (rs.Fields("subject"))
            rs.MoveNext
        Loop
    End If
End Sub


Private Sub Combo4_Click()
If Combo3.Text = "(Select Sem)" Then
    MsgBox "Please select semester first", vbCritical + vbOKOnly, "Warning"
Else
    List1.Clear
    If Combo4.Text = "(Select)" Then
            MsgBox "Please select the package or subject", vbCritical + vbOKOnly, "Warning"
    ElseIf Combo4.Text = "All Package" Then
        Command3.Enabled = False
        List1.Visible = True
        Combo5.Enabled = False
        Set rs = cnn.Execute("select * from package_details where field='" + Combo1.Text + "' and sem='" + Combo3.Text + "' and package_type='" + Combo4.Text + "'")
        Dim i As Integer
        i = 4
        If (Not rs.EOF) Then
                Label12.Caption = rs.Fields("price")
                While (i < 10)
                    If rs.Fields(i) <> "Null" Then
                        List1.AddItem (rs.Fields(i))
                    End If
                    i = i + 1
                Wend
        Else
            MsgBox "All Package is not available for selected semester", vbInformation + vbOKOnly, "Information"
        End If
    ElseIf Combo4.Text = "Small Package" Then
        Command3.Enabled = False
        List1.Visible = True
        Combo5.Enabled = False
        Set rs = cnn.Execute("select * from package_details where field='" + Combo1.Text + "' and sem='" + Combo3.Text + "' and package_type='" + Combo4.Text + "'")
        Dim j As Integer
        j = 4
        If (Not rs.EOF) Then
                Label12.Caption = rs.Fields("price")
                While (j < 10)
                    If rs.Fields(j) <> "Null" Then
                        List1.AddItem (rs.Fields(j))
                    End If
                    j = j + 1
                Wend
        Else
            MsgBox "Small Package is not available for selected semester", vbInformation + vbOKOnly, "Information"
        End If
    Else
            Command3.Enabled = True
            Label12.Caption = 0
            Combo5.Enabled = True
            List1.Visible = True
    End If
End If
End Sub

Private Sub Combo5_Click()
If Combo1.Text <> "OTHER" Then
    If List1.ListCount = 2 Then
        MsgBox "Sorry two subjects are already selected", vbExclamation + vbOKOnly, "Information"
    ElseIf List1.ListCount = 0 Then
            List1.AddItem (Combo5.Text)
            Set rs = cnn.Execute("select * from subject_details where years='" + Combo1.Text + "' and sem='" + Combo3.Text + "'and subject='" + Combo5.Text + "'")
            If (Not rs.EOF) Then
                Dim a As Single
                a = rs.Fields("cost")
                Label12.Caption = Val(Label12.Caption) + a
                Label16.Caption = a
            End If
    Else
        For m = 0 To List1.ListCount
            If List1.List(m) <> Combo5.Text Then
                If m = List1.ListCount Then
                    List1.AddItem (Combo5.Text)
                    Set rs = cnn.Execute("select * from subject_details where years='" + Combo1.Text + "' and sem='" + Combo3.Text + "'and subject='" + Combo5.Text + "'")
                    If (Not rs.EOF) Then
                        Dim a1 As Single
                        a1 = rs.Fields("cost")
                        Label12.Caption = Val(Label12.Caption) + a1
                        Label15.Caption = a1
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
End Sub


Private Sub Combo5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command1_Click()
If Len(Text2.Text) < 10 Then
    MsgBox "Mobile numbers should be of 10 digits", vbCritical + vbOKOnly, "Warning"
ElseIf Len(Text5.Text) < 6 Then
    MsgBox "Pincode should be of 6 digits", vbOKOnly + vbExclamation, "Warning"
Else
    If Combo1.Text = "(Select)" Then
        MsgBox "Please select the field ", vbCritical + vbOKOnly, "Warning"
    ElseIf Text1.Text = "" Then
        MsgBox "Please Enter your Name ", vbCritical + vbOKOnly, "Warning"
    ElseIf Text2.Text = "" Then
        MsgBox "Please Enter your personal number", vbCritical + vbOKOnly, "Warning"
    ElseIf Text6.Text = "" Then
        MsgBox "Please Enter your parents mobile number", vbCritical + vbOKOnly, "Warning"
    ElseIf Combo2.Text = "(Select College)" Then
        MsgBox "Please select college of student", vbCritical + vbOKOnly, "Warning"
    ElseIf Combo3.Text = "(Select Sem)" Then
        MsgBox "Please select the semester of the student", vbCritical + vbOKOnly, "Warning"
    ElseIf Combo4.Text = "(Select)" Then
        MsgBox "Please select the package", vbCritical + vbOKOnly, "Warning"
    ElseIf List1.ListCount = 0 Then
        MsgBox "Please select the subject", vbCritical + vbOKOnly, "Warning"
    ElseIf Combo6.Text = "(Select)" Then
        MsgBox "Please select the season", vbCritical + vbOKOnly, "Warning"
    Else
        s = "insert into admission (srno,rollno,sname,selfmobno,parentmobno,field,college,subject,admitdate,totalfee,sem,area,landmark,pincode,season) values ('" & Label2.Caption & "','" & Label5.Caption & "','" & Text1.Text & "','" & Text2.Text & "','" & Text6.Text & "','" & Combo1.Text & "','" & Combo2.Text & "','" & Combo4.Text & "','" & DTPicker1.Value & "','" & Label12.Caption & "','" & Combo3.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Combo6.Text & "')"
        cnn.Execute s
        MsgBox "Student details saved successfully", vbInformation + vbOKOnly, "Successfull"
        If List1.ListCount = 1 Then
                m = "insert into stud_sub_detail (sname,rollno,sem,package_type,sub1,sub2,sub3,sub4,sub5,sub6,field,season) values ('" & Text1.Text & "','" & Label5.Caption & "','" & Combo3.Text & "','" & Combo4.Text & "','" & List1.List(0) & "','" & "Null" & "','" & "Null" & "','" & "Null" & "','" & "Null" & "','" & "Null" & "','" & Combo1.Text & "','" & Combo6.Text & "')"
                cnn.Execute m
        ElseIf List1.ListCount = 2 Then
                m = "insert into stud_sub_detail (sname,rollno,sem,package_type,sub1,sub2,sub3,sub4,sub5,sub6,field,season) values ('" & Text1.Text & "','" & Label5.Caption & "','" & Combo3.Text & "','" & Combo4.Text & "','" & List1.List(0) & "','" & List1.List(1) & "','" & "Null" & "','" & "Null" & "','" & "Null" & "','" & "Null" & "','" & Combo1.Text & "','" & Combo6.Text & "')"
                cnn.Execute m
        ElseIf List1.ListCount = 3 Then
                m = "insert into stud_sub_detail (sname,rollno,sem,package_type,sub1,sub2,sub3,sub4,sub5,sub6,field,season) values ('" & Text1.Text & "','" & Label5.Caption & "','" & Combo3.Text & "','" & Combo4.Text & "','" & List1.List(0) & "','" & List1.List(1) & "','" & List1.List(2) & "','" & "Null" & "','" & "Null" & "','" & "Null" & "','" & Combo1.Text & "','" & Combo6.Text & "')"
                cnn.Execute m
        ElseIf List1.ListCount = 4 Then
                m = "insert into stud_sub_detail (sname,rollno,sem,package_type,sub1,sub2,sub3,sub4,sub5,sub6,field,season) values ('" & Text1.Text & "','" & Label5.Caption & "','" & Combo3.Text & "','" & Combo4.Text & "','" & List1.List(0) & "','" & List1.List(1) & "','" & List1.List(2) & "','" & List1.List(3) & "','" & "Null" & "','" & "Null" & "','" & Combo1.Text & "','" & Combo6.Text & "')"
                cnn.Execute m
        ElseIf List1.ListCount = 5 Then
                m = "insert into stud_sub_detail (sname,rollno,sem,package_type,sub1,sub2,sub3,sub4,sub5,sub6,field,season) values ('" & Text1.Text & "','" & Label5.Caption & "','" & Combo3.Text & "','" & Combo4.Text & "','" & List1.List(0) & "','" & List1.List(1) & "','" & List1.List(2) & "','" & List1.List(3) & "','" & List1.List(4) & "','" & "Null" & "','" & Combo1.Text & "','" & Combo6.Text & "')"
                cnn.Execute m
        ElseIf List1.ListCount = 6 Then
                m = "insert into stud_sub_detail (sname,rollno,sem,package_type,sub1,sub2,sub3,sub4,sub5,sub6,field,season) values ('" & Text1.Text & "','" & Label5.Caption & "','" & Combo3.Text & "','" & Combo4.Text & "','" & List1.List(0) & "','" & List1.List(1) & "','" & List1.List(2) & "','" & List1.List(3) & "','" & List1.List(4) & "','" & List1.List(5) & "','" & Combo1.Text & "','" & Combo6.Text & "')"
                cnn.Execute m
        End If
        Command2.Enabled = True
        Command1.Enabled = False
        
    End If
End If
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command2_Click()
    If Combo4.Text = "Select Subjects" Then
        If List1.ListCount = 1 Then
            count5 = 1
            price1 = Label12.Caption
        ElseIf List1.ListCount = 2 Then
            count5 = 2
            If Val(Label16.Caption) > Val(Label15.Caption) Then
                price1 = Label16.Caption
                price2 = Label15.Caption
            Else
                price1 = Label15.Caption
                price2 = Label16.Caption
            End If
        End If
    End If
        date1 = DTPicker1.Value
     rollno1 = Label5.Caption
     sname1 = Text1.Text
     sem = Combo3.Text
     package_type = Combo4.Text
     field = Combo1.Text
     cost = Val(Label12.Caption)
     frm_transaction.Show
     Unload Me
End Sub

Private Sub Command3_Click()
    Dim a As Integer
    For i = List1.ListCount - 1 To 0 Step -1
        If List1.Selected(i) = True Then
            Set rs = cnn.Execute("select * from subject_details where years='" + Combo1.Text + "' and sem='" + Combo3.Text + "'and subject='" + List1.List(i) + "'")
            If (Not rs.EOF) Then
                a = Val(rs.Fields("cost"))
            End If
            Label12.Caption = Val(Label12.Caption) - a
            List1.RemoveItem (i)
        End If
    Next i
    
    
    
End Sub

Private Sub Command5_Click()
    If Text9.Text <> "" Then
    If List1.ListCount = 2 Then
            MsgBox "Sorry two subjects are already selected", vbExclamation + vbOKOnly, "Information"
    ElseIf List1.ListCount = 0 Then
        List1.AddItem (Combo5.Text)
        Label12.Caption = Val(Label12.Caption) + Val(Text9.Text)
        Label15.Caption = Text9.Text
        Text9.Text = ""
    Else
        For m = 0 To List1.ListCount
            If List1.List(m) <> Combo5.Text Then
                If m = List1.ListCount Then
                    List1.AddItem (Combo5.Text)
                    Label12.Caption = Val(Label12.Caption) + Val(Text9.Text)
                    Label16.Caption = Text9.Text
                    Text9.Text = ""
                    Exit For
                End If
            Else
                MsgBox "Subject already selected", vbInformation + vbOKOnly, "Information"
                Exit For
            End If
        Next
    End If
Else
    MsgBox "Please enter the price for the selected subject", vbExclamation + vbOKOnly, "Information"
End If
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
     Me.Top = 600
    Me.Left = 1000
    count1 = 0
    count2 = 1
    connect
    Set rs = cnn.Execute("select * from admission")
    Do While Not rs.EOF
        a = rs.Fields("srno")
        rs.MoveNext
    Loop
    Label2.Caption = Val(a) + 1
    Set rs = cnn.Execute("select * from college'")
    Do While Not rs.EOF
        Combo2.AddItem (rs.Fields("cname"))
        rs.MoveNext
    Loop
    a = Date
    b = Mid(a, 9, 10)
    Combo6.AddItem ("Winter " + b)
    Combo6.AddItem ("Summer " + b)
End Sub



Private Sub Text1_Change()
    Text2.Enabled = True
End Sub

Private Sub Text2_Change()
    Text6.Enabled = True
End Sub

Private Sub Text2_LostFocus()
    If Len(Text2.Text) < 10 Then
        MsgBox "Mobile numbers should be of 10 digits", vbCritical + vbOKOnly, "Warning"
    End If
End Sub

Private Sub Text5_LostFocus()
If Len(Text5.Text) < 6 Then
    MsgBox "Pincode should be of 6 digits", vbOKOnly + vbExclamation, "Warning"
    Text5.SetFocus
End If
End Sub

Private Sub Text6_Change()
    Combo2.Enabled = True
    Combo3.Enabled = True
    Combo4.Enabled = True
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub
