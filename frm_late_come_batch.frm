VERSION 5.00
Begin VB.Form frm_late_come_batch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Student to batch (Late Admission) :"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List2 
      Height          =   4905
      Left            =   8520
      TabIndex        =   24
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   4905
      Left            =   7200
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   600
      Width           =   5055
      Begin VB.ComboBox Combo4 
         Height          =   405
         Left            =   840
         TabIndex        =   21
         Text            =   "(Batch)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frm_late_come_batch.frx":0000
         Left            =   2880
         List            =   "frm_late_come_batch.frx":001C
         TabIndex        =   20
         Text            =   "(Sem)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox Combo5 
         Height          =   405
         ItemData        =   "frm_late_come_batch.frx":0062
         Left            =   120
         List            =   "frm_late_come_batch.frx":0072
         TabIndex        =   19
         Text            =   "(Field)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   4455
      Begin VB.OptionButton Option2 
         Caption         =   "By Batch"
         Height          =   285
         Left            =   2640
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Semesters"
         Height          =   285
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   0
      TabIndex        =   10
      Top             =   3360
      Width           =   6855
      Begin VB.CommandButton Command2 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         Left            =   2520
         TabIndex        =   5
         Text            =   "(Select Batch)"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   2520
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   2520
         TabIndex        =   3
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Select Batch :"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Field :"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Student Roll No :"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Student Name :"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student Details :"
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   6855
      Begin VB.ComboBox Combo6 
         Height          =   405
         Left            =   3240
         TabIndex        =   22
         Text            =   "(Select Name)"
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox Combo3 
         Height          =   405
         Left            =   3240
         TabIndex        =   0
         Text            =   "(Select)"
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5520
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Select Season :"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Enter RollNo or Name :"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frm_late_come_batch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim count1 As Integer
Dim count2 As Integer



Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command2_Click
    End If
End Sub

Private Sub Combo2_Click()
    Combo6.Clear
    If Combo5.Text = "(Field)" Then
        MsgBox "Please select the Field", vbQuestion + vbOKOnly, "Warning"
    ElseIf Combo2.Text = "(Sem)" Then
        MsgBox "Please select the Semester", vbQuestion + vbOKOnly, "Warning"
    Else
        Set rs = cnn.Execute("select * from admission where field='" & Combo5.Text & "' and sem='" & Combo2.Text & "'")
        Do While Not rs.EOF
            Combo6.AddItem (rs.Fields("sname"))
            rs.MoveNext
        Loop
    End If
End Sub


Private Sub Combo4_Click()
            Dim flag As Integer
            flag = 0
            List1.Clear
            List2.Clear
            Set rs1 = cnn.Execute("select sname,rollno from stud_sub_detail where sub1='" & Combo4.Text & "'or sub2='" & Combo4.Text & "'or sub3='" & Combo4.Text & "' or sub4='" & Combo4.Text & "'or sub5='" & Combo4.Text & "'or sub6='" & Combo4.Text & "'")
            If (Not rs1.EOF) Then
                Do While Not rs1.EOF
                                List1.AddItem (rs1.Fields("sname"))
                                List2.AddItem (rs1.Fields("rollno"))
                                rs1.MoveNext
                Loop
            Else
                MsgBox "Students not found", vbExclamation + vbOKOnly, "Information"
            End If
            If List1.ListCount <> 0 Then
                Set rs = cnn.Execute("select * from stud_batch_details where batch='" & Combo4.Text & "'")
                For i = 0 To List1.ListCount - 1
                    Do While Not rs.EOF
                        If rs.Fields("sname") = List1.List(i) Then
                            flag = 1
                            Exit Do
                        End If
                        rs.MoveNext
                    Loop
                    If flag = 0 Then
                        Combo6.AddItem (List1.List(i))
                        rs.MoveFirst
                    Else
                        flag = 0
                        rs.MoveFirst
                    End If
                Next
            End If
    
    
    
    
    
    
    
    
    
    'Combo6.Clear
    'Set rs = cnn.Execute("Select * from stud_batch_details where batch='" & Combo4.Text & "'")
    'Do While Not rs.EOF
      '  Combo6.AddItem (rs.Fields("sname"))
       ' rs.MoveNext
    'Loop
End Sub

Private Sub Combo5_Click()
    If Combo5.Text = "OTHER" Then
        If count2 = 0 And Combo5.Text = "OTHER" Then
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

    Combo6.Clear
    If Combo2.Text <> "(Sem)" Then
        Set rs = cnn.Execute("select * from admission where field='" & Combo5.Text & "' and sem='" & Combo2.Text & "'")
        Do While Not rs.EOF
            Combo6.AddItem (rs.Fields("sname"))
            rs.MoveNext
        Loop
    End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command1_Click()
    If Combo6.Text = "(Select Name)" Then
        MsgBox "Please select the name of the student", vbExclamation + vbOKOnly, "Warning"
    ElseIf Combo3.Text = "(Select)" Then
        MsgBox "Please select season", vbExclamation + vbOKOnly, "Warning"
    Else
        Set rs = cnn.Execute("Select * from admission where sname='" & Combo6.Text & "' and season='" & Combo3.Text & "'")
        If (Not rs.EOF) Then
            Text2.Text = rs.Fields("sname")
            Text3.Text = rs.Fields("rollno")
            Text4.Text = rs.Fields("field")
            Command2.Enabled = True
        Else
            MsgBox "Student not found", vbExclamation + vbOKOnly, "Warning"
        End If
    End If
End Sub

Private Sub Command2_Click()
    If Combo1.Text = "(Select Batch)" Then
        MsgBox "Please select the batch", vbExclamation + vbOKOnly, "Warning"
    Else
        s = "insert into stud_batch_details (rollno,sname,batch,field,season) values ('" & Text3.Text & "','" & Text2.Text & "','" & Combo1.Text & "','" & Text4.Text & "','" & Combo3.Text & "')"
        cnn.Execute s
        MsgBox "Student added successfully to the batch", vbInformation + vbOKOnly, "Information"
        Unload Me
    End If
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command2_Click
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 1000
    Me.Left = 3500
    connect
    a = Date
    b = Mid(a, 9, 10)
    Combo3.AddItem ("Winter " + b)
    Combo3.AddItem ("Summer " + b)
    Set rs = cnn.Execute("select distinct  batch from batch_details order by batch")
    Do While Not rs.EOF
        Combo1.AddItem (rs.Fields("batch"))
        rs.MoveNext
    Loop
    count1 = 0
    count2 = 0
     Set rs = cnn.Execute("select * from batch_details")
    If (Not rs.EOF) Then
        Do While Not rs.EOF
            Combo4.AddItem (rs.Fields("batch"))
            rs.MoveNext
        Loop
    End If
End Sub

Private Sub Option1_Click()
     Combo5.Visible = True
    Combo2.Visible = True
    Combo4.Visible = False
End Sub

Private Sub Option2_Click()
     Combo5.Visible = False
    Combo2.Visible = False
    Combo4.Visible = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub
