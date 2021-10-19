VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Old Student Admission :"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13050
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   13050
   Begin VB.CommandButton Command4 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   11160
      TabIndex        =   43
      Top             =   7800
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   4335
      Left            =   10560
      TabIndex        =   36
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   495
      Left            =   11760
      TabIndex        =   30
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   9120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   9120
      TabIndex        =   28
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox Combo4 
      Height          =   405
      ItemData        =   "frmAdmission.frx":0000
      Left            =   3600
      List            =   "frmAdmission.frx":000A
      TabIndex        =   0
      Text            =   "(Select)"
      Top             =   240
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Student Detail :"
      Height          =   7815
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   10095
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   5880
         TabIndex        =   49
         Top             =   4680
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton Command5 
            Caption         =   "Add"
            Height          =   375
            Left            =   1200
            TabIndex        =   51
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox Text9 
            Height          =   405
            Left            =   1800
            TabIndex        =   10
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label26 
            Caption         =   "Enter Price :"
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.ComboBox Combo7 
         Height          =   405
         Left            =   7680
         TabIndex        =   4
         Text            =   "(Select)"
         Top             =   960
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Caption         =   "Address Details :"
         Height          =   2175
         Left            =   5880
         TabIndex        =   44
         Top             =   2040
         Width           =   4095
         Begin VB.TextBox Text8 
            Height          =   405
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   17
            Top             =   1440
            Width           =   2175
         End
         Begin VB.TextBox Text7 
            Height          =   405
            Left            =   1800
            TabIndex        =   16
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox Text3 
            Height          =   405
            Left            =   1800
            TabIndex        =   15
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label24 
            Caption         =   "Pincode :"
            Height          =   375
            Left            =   120
            TabIndex        =   47
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "Land-Mark :"
            Height          =   375
            Left            =   120
            TabIndex        =   46
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "Area :"
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.ComboBox Combo6 
         Enabled         =   0   'False
         Height          =   405
         Left            =   3360
         TabIndex        =   14
         Text            =   "(Select Subjects)"
         Top             =   6600
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   6600
         TabIndex        =   3
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   159514625
         CurrentDate     =   43858
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frmAdmission.frx":0032
         Left            =   3360
         List            =   "frmAdmission.frx":003F
         TabIndex        =   13
         Text            =   "(Select Package)"
         Top             =   6000
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Proceed"
         Enabled         =   0   'False
         Height          =   495
         Left            =   8520
         TabIndex        =   19
         Top             =   7080
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Add Details"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6600
         TabIndex        =   18
         Top             =   7080
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   3360
         TabIndex        =   6
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ComboBox Combo5 
         Height          =   405
         ItemData        =   "frmAdmission.frx":0070
         Left            =   3360
         List            =   "frmAdmission.frx":0080
         TabIndex        =   9
         Text            =   "(Select Field)"
         Top             =   4680
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   3360
         TabIndex        =   7
         Top             =   3360
         Width           =   2295
      End
      Begin VB.ComboBox Combo3 
         Height          =   405
         ItemData        =   "frmAdmission.frx":00A0
         Left            =   3360
         List            =   "frmAdmission.frx":00BC
         TabIndex        =   12
         Text            =   "(Select Sem)"
         Top             =   5280
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frmAdmission.frx":0102
         Left            =   3360
         List            =   "frmAdmission.frx":0104
         TabIndex        =   8
         Text            =   "(Select College)"
         Top             =   4080
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   3360
         TabIndex        =   5
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label25 
         Caption         =   "Select Season :"
         Height          =   375
         Left            =   5640
         TabIndex        =   48
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label21 
         Caption         =   "Select Subjects :"
         Height          =   375
         Left            =   360
         TabIndex        =   42
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label Label20 
         Caption         =   "Date :"
         Height          =   375
         Left            =   5640
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Select Package :"
         Height          =   375
         Left            =   360
         TabIndex        =   40
         Top             =   6120
         Width           =   2295
      End
      Begin VB.Label Label15 
         Caption         =   "New Roll No :"
         Height          =   375
         Left            =   360
         TabIndex        =   35
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   2040
         TabIndex        =   34
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   1440
         TabIndex        =   33
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Sr No :"
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Select Field :"
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Select Semester :"
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   5400
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Select College :"
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Parents Mobile No :"
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Personal Mobile No :"
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Student Name :"
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label5 
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Old Roll No :"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Label Label18 
      Caption         =   "0"
      Height          =   375
      Left            =   10800
      TabIndex        =   39
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "0"
      Height          =   375
      Left            =   10800
      TabIndex        =   38
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label16 
      Caption         =   "0"
      Height          =   375
      Left            =   10800
      TabIndex        =   37
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "Enter Student Rollno :"
      Height          =   255
      Left            =   6240
      TabIndex        =   29
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label12 
      Caption         =   "Enter Student Name :"
      Height          =   255
      Left            =   6120
      TabIndex        =   27
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label11 
      Caption         =   "Search  Student By :"
      Height          =   375
      Left            =   600
      TabIndex        =   26
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim count5 As Integer
Dim count1 As Integer
Dim count2 As Integer
Private Sub Combo1_Click()
If Combo3.Text = "(Select Sem)" Then
    MsgBox "Please select semester first", vbCritical + vbOKOnly, "Warning"
Else
    List1.Clear
    If Combo1.Text = "(Select Package)" Then
            MsgBox "Please select the package or subject", vbCritical + vbOKOnly, "Warning"
    ElseIf Combo1.Text = "All Package" Then
        Command4.Enabled = False
        List1.Visible = True
        Combo6.Enabled = False
        Set rs = cnn.Execute("select * from package_details where field='" + Combo5.Text + "' and sem='" + Combo3.Text + "' and package_type='" + Combo1.Text + "'")
        Dim i As Integer
        i = 4
        If (Not rs.EOF) Then
                Label16.Caption = rs.Fields("price")
                While (i < 10)
                    If rs.Fields(i) <> "Null" Then
                        List1.AddItem (rs.Fields(i))
                    End If
                    i = i + 1
                Wend
        Else
            MsgBox "All Package is not available for selected semester", vbInformation + vbOKOnly, "Information"
        End If
    ElseIf Combo1.Text = "Small Package" Then
        Command4.Enabled = False
        List1.Visible = True
        Combo6.Enabled = False
        Set rs = cnn.Execute("select * from package_details where field='" + Combo5.Text + "' and sem='" + Combo3.Text + "' and package_type='" + Combo1.Text + "'")
        Dim j As Integer
        j = 4
        If (Not rs.EOF) Then
                Label16.Caption = rs.Fields("price")
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
            Command4.Enabled = True
            Label16.Caption = 0
            Combo6.Enabled = True
            List1.Visible = True
    End If
End If
End Sub

Private Sub Combo3_Click()
    Combo6.Clear
    List1.Clear
    Label16.Caption = 0
    Label17.Caption = 0
    Label18.Caption = 0
    If Combo5.Text = "OTHER" Then
        Set rs = cnn.Execute("select distinct subject from subject_details order by subject")
        Do While Not rs.EOF
            Combo6.AddItem (rs.Fields("subject"))
            rs.MoveNext
        Loop
    Else
        Set rs = cnn.Execute("Select * from subject_details where years='" + Combo5.Text + "' and sem='" + Combo3.Text + "'")
        Do While Not rs.EOF
            Combo6.AddItem (rs.Fields("subject"))
            rs.MoveNext
        Loop
    End If
End Sub

Private Sub Combo4_Click()
    If Combo4.Text = "(Select)" Then
        MsgBox "Please select the option to search", vbCritical + vbOKOnly, "warning"
    ElseIf Combo4.Text = "By Student Name" Then
        Label12.Visible = True
        Text4.Visible = True
        Label13.Visible = False
        Text5.Visible = False
        Command1.Visible = True
    ElseIf Combo4.Text = "By Student Rollno" Then
        Label12.Visible = False
        Text4.Visible = False
        Label13.Visible = True
        Text5.Visible = True
         Command1.Visible = True
    End If
End Sub
Private Sub Combo5_Click()
    If Combo5.Text = "OTHER" Then
        Frame3.Visible = True
    Else
        Frame3.Visible = False
    End If
    If Combo5.Text = "OTHER" Then
        If count2 = 0 And Combo5.Text = "OTHER" Then
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

Private Sub Combo6_Click()
If Combo5.Text <> "OTHER" Then
    If List1.ListCount = 2 Then
        MsgBox "Sorry two subjects are already selected", vbExclamation + vbOKOnly, "Information"
    ElseIf List1.ListCount = 0 Then
            List1.AddItem (Combo6.Text)
            Set rs = cnn.Execute("select * from subject_details where years='" + Combo5.Text + "' and sem='" + Combo3.Text + "'and subject='" + Combo6.Text + "'")
            If (Not rs.EOF) Then
                Dim a As Single
                a = rs.Fields("cost")
                Label16.Caption = Val(Label16.Caption) + a
                Label17.Caption = a
            End If
    Else
        For m = 0 To List1.ListCount
            If List1.List(m) <> Combo6.Text Then
                If m = List1.ListCount Then
                    List1.AddItem (Combo6.Text)
                    Set rs = cnn.Execute("select * from subject_details where years='" + Combo5.Text + "' and sem='" + Combo3.Text + "'and subject='" + Combo6.Text + "'")
                    If (Not rs.EOF) Then
                        Dim a1 As Single
                        a1 = rs.Fields("cost")
                        Label16.Caption = Val(Label16.Caption) + a1
                        Label18.Caption = a1
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


Private Sub Command1_Click()
    Set rs = cnn.Execute("select * from admission")
    Do While Not rs.EOF
        k = rs.Fields("srno")
        rs.MoveNext
    Loop
    Label2.Caption = Val(k) + 1
            Dim a As String
            Dim b As String
            Dim c As String
            Dim d As String
    If Combo4.Text = "By Student Name" Then
        Set rs = cnn.Execute("select * from admission where sname='" + Text4.Text + "'")
        If (Not rs.EOF) Then
            Combo5.Text = rs.Fields("field")
            If Combo5.Text = "OTHER" Then
                Frame3.Visible = True
            End If
            Label5.Caption = rs.Fields("rollno")
            Text1.Text = rs.Fields("sname")
            Text2.Text = rs.Fields("selfmobno")
            Text6.Text = rs.Fields("parentmobno")
            Combo2.Text = rs.Fields("college")
            Text3.Text = rs.Fields("area")
            Text7.Text = rs.Fields("landmark")
            Text8.Text = rs.Fields("pincode")
            a = rs.Fields("rollno")
            If rs.Fields("field") = "OTHER" Then
                If count2 = 0 And Combo5.Text = "OTHER" Then
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
                                                             
            If rs.Fields("field") = "BSC-IT" Then
                b = Mid(a, 4)
                c = Val(b) + 1
                d = Mid(a, 1, 3)
                Label3.Caption = d + c
            Else
                b = Mid(a, 3)
                c = Val(b) + 1
                d = Mid(a, 1, 2)
                Label3.Caption = d + c
            End If
            Command2.Enabled = True
        Else
            MsgBox "No student available with this name", vbCritical + vbOKOnly, "warning"
        End If
    ElseIf Combo4.Text = "By Student Rollno" Then
        Set rs = cnn.Execute("select * from admission where rollno='" + Text5.Text + "'")
        If (Not rs.EOF) Then
            Combo5.Text = rs.Fields("field")
            Label5.Caption = rs.Fields("rollno")
            Text1.Text = rs.Fields("sname")
            Text2.Text = rs.Fields("selfmobno")
            Text6.Text = rs.Fields("parentmobno")
            Combo2.Text = rs.Fields("college")
            a = rs.Fields("rollno")
            If rs.Fields("field") = "BSC-IT" Then
                b = Mid(a, 4)
                c = Val(b) + 1
                d = Mid(a, 1, 3)
                Label3.Caption = d + c
            Else
                b = Mid(a, 3)
                c = Val(b) + 1
                d = Mid(a, 1, 2)
                Label3.Caption = d + c
            End If
            
            If rs.Fields("field") = "OTHER" Then
                If count2 = 0 And Combo5.Text = "OTHER" Then
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
            
            Command2.Enabled = True
        Else
            MsgBox "No student available with this rollno", vbCritical + vbOKOnly, "warning"
            count10 = 0
        End If
            
    End If
End Sub

Private Sub Command2_Click()
If Len(Text2.Text) < 10 Then
    MsgBox "Mobile numbers should be of 10 digits", vbCritical + vbOKOnly, "Warning"
ElseIf Len(Text8.Text) < 6 Then
    MsgBox "Pincode should be of 6 digits", vbOKOnly + vbExclamation, "Warning"
Else
    If Combo5.Text = "(Select Field)" Then
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
    ElseIf Combo1.Text = "(Select)" Then
        MsgBox "Please select the package", vbCritical + vbOKOnly, "Warning"
    ElseIf List1.ListCount = 0 Then
        MsgBox "Please select the subject", vbCritical + vbOKOnly, "Warning"
    Else
        s = "insert into admission (srno,rollno,sname,selfmobno,parentmobno,field,college,subject,admitdate,totalfee,sem,area,landmark,pincode,season) values ('" & Label2.Caption & "','" & Label3.Caption & "','" & Text1.Text & "','" & Text2.Text & "','" & Text6.Text & "','" & Combo5.Text & "','" & Combo2.Text & "','" & Combo1.Text & "','" & DTPicker1.Value & "','" & Label16.Caption & "','" & Combo3.Text & "','" & Text3.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Combo7.Text & "')"
        cnn.Execute s
        MsgBox "Student details saved successfully", vbInformation + vbOKOnly, "Successfull"
        If List1.ListCount = 1 Then
                m = "insert into stud_sub_detail (sname,rollno,sem,package_type,sub1,sub2,sub3,sub4,sub5,sub6,field,season) values ('" & Text1.Text & "','" & Label3.Caption & "','" & Combo3.Text & "','" & Combo1.Text & "','" & List1.List(0) & "','" & "Null" & "','" & "Null" & "','" & "Null" & "','" & "Null" & "','" & "Null" & "','" & Combo5.Text & "','" & Combo7.Text & "')"
                cnn.Execute m
        ElseIf List1.ListCount = 2 Then
                m = "insert into stud_sub_detail (sname,rollno,sem,package_type,sub1,sub2,sub3,sub4,sub5,sub6,field,season) values ('" & Text1.Text & "','" & Label3.Caption & "','" & Combo3.Text & "','" & Combo1.Text & "','" & List1.List(0) & "','" & List1.List(1) & "','" & "Null" & "','" & "Null" & "','" & "Null" & "','" & "Null" & "','" & Combo5.Text & "','" & Combo7.Text & "')"
                cnn.Execute m
        ElseIf List1.ListCount = 3 Then
                m = "insert into stud_sub_detail (sname,rollno,sem,package_type,sub1,sub2,sub3,sub4,sub5,sub6,field,season) values ('" & Text1.Text & "','" & Label3.Caption & "','" & Combo3.Text & "','" & Combo1.Text & "','" & List1.List(0) & "','" & List1.List(1) & "','" & List1.List(2) & "','" & "Null" & "','" & "Null" & "','" & "Null" & "','" & Combo5.Text & "','" & Combo7.Text & "')"
                cnn.Execute m
        ElseIf List1.ListCount = 4 Then
                m = "insert into stud_sub_detail (sname,rollno,sem,package_type,sub1,sub2,sub3,sub4,sub5,sub6,field,season) values ('" & Text1.Text & "','" & Label3.Caption & "','" & Combo3.Text & "','" & Combo1.Text & "','" & List1.List(0) & "','" & List1.List(1) & "','" & List1.List(2) & "','" & List1.List(3) & "','" & "Null" & "','" & "Null" & "','" & Combo5.Text & "','" & Combo7.Text & "')"
                cnn.Execute m
        ElseIf List1.ListCount = 5 Then
                m = "insert into stud_sub_detail (sname,rollno,sem,package_type,sub1,sub2,sub3,sub4,sub5,sub6,field,season) values ('" & Text1.Text & "','" & Label3.Caption & "','" & Combo3.Text & "','" & Combo1.Text & "','" & List1.List(0) & "','" & List1.List(1) & "','" & List1.List(2) & "','" & List1.List(3) & "','" & List1.List(4) & "','" & "Null" & "','" & Combo5.Text & "','" & Combo7.Text & "')"
                cnn.Execute m
        ElseIf List1.ListCount = 6 Then
                m = "insert into stud_sub_detail (sname,rollno,sem,package_type,sub1,sub2,sub3,sub4,sub5,sub6,field,season) values ('" & Text1.Text & "','" & Label3.Caption & "','" & Combo3.Text & "','" & Combo1.Text & "','" & List1.List(0) & "','" & List1.List(1) & "','" & List1.List(2) & "','" & List1.List(3) & "','" & List1.List(4) & "','" & List1.List(5) & "','" & Combo5.Text & "','" & Combo7.Text & "')"
                cnn.Execute m
        End If
        Command3.Enabled = True
        Command2.Enabled = False
    End If
End If
End Sub

Private Sub Command3_Click()
If Combo1.Text = "Select Subjects" Then
        If List1.ListCount = 1 Then
            count1 = 1
            price1 = Label16.Caption
        ElseIf List1.ListCount = 2 Then
            count1 = 2
            If Val(Label17.Caption) > Val(Label18.Caption) Then
                price1 = Label17.Caption
                price2 = Label18.Caption
            Else
                price1 = Label18.Caption
                price2 = Label17.Caption
            End If
        End If
    End If
    date1 = DTPicker1.Value
     rollno1 = Label3.Caption
     sname1 = Text1.Text
     sem = Combo3.Text
     package_type = Combo1.Text
     field = Combo5.Text
     cost = Val(Label16.Caption)
     frm_transaction.Show
     Unload Me
End Sub

Private Sub Command4_Click()
     Dim a As Integer
    For i = List1.ListCount - 1 To 0 Step -1
        If List1.Selected(i) = True Then
            Set rs = cnn.Execute("select * from subject_details where years='" + Combo5.Text + "' and sem='" + Combo3.Text + "'and subject='" + List1.List(i) + "'")
            If (Not rs.EOF) Then
                a = Val(rs.Fields("cost"))
            End If
            Label16.Caption = Val(Label16.Caption) - a
            List1.RemoveItem (i)
        End If
    Next i
    
End Sub

Private Sub Command5_Click()
If Text9.Text <> "" Then
    If List1.ListCount = 2 Then
            MsgBox "Sorry two subjects are already selected", vbExclamation + vbOKOnly, "Information"
    ElseIf List1.ListCount = 0 Then
        List1.AddItem (Combo6.Text)
        Label16.Caption = Val(Label16.Caption) + Val(Text9.Text)
        Label17.Caption = Text9.Text
        Text9.Text = ""
    Else
        For m = 0 To List1.ListCount
            If List1.List(m) <> Combo6.Text Then
                If m = List1.ListCount Then
                    List1.AddItem (Combo6.Text)
                    Label16.Caption = Val(Label16.Caption) + Val(Text9.Text)
                    Label18.Caption = Text9.Text
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
    Me.Top = 400
    Me.Left = 800
    count1 = 0
    count2 = 1
    Dim count5 As Integer
    count5 = 0
    'Me.Move (MDIForm1.ScaleWidth - Me.Width)
    connect
    Set rs = cnn.Execute("select * from college'")
    Do While Not rs.EOF
        Combo2.AddItem (rs.Fields("cname"))
        rs.MoveNext
    Loop
    a = Date
    b = Mid(a, 9, 10)
    Combo7.AddItem ("Winter " + b)
    Combo7.AddItem ("Summer " + b)
End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
   
        If KeyAscii = 13 Then
            Call Command1_Click
           
        End If
    
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    
        If KeyAscii = 13 Then
            Call Command1_Click
            
        End If
    
End Sub

