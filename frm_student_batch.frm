VERSION 5.00
Begin VB.Form frm_student_batch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Students to Batch :"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10110
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
   ScaleHeight     =   4605
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List3 
      Height          =   4335
      Left            =   7440
      TabIndex        =   13
      Top             =   120
      Width           =   2535
   End
   Begin VB.ListBox List2 
      Height          =   7185
      Left            =   9960
      TabIndex        =   11
      Top             =   -2760
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.ListBox List1 
      Height          =   4335
      Left            =   10080
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Students Batch :"
      Height          =   4575
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   7335
      Begin VB.ComboBox Combo4 
         Height          =   405
         Left            =   2160
         TabIndex        =   1
         Text            =   "(Select)"
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Add"
         Height          =   495
         Left            =   3360
         TabIndex        =   5
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Search"
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   3480
         Width           =   1695
      End
      Begin VB.ComboBox Combo3 
         Height          =   405
         ItemData        =   "frm_student_batch.frx":0000
         Left            =   2160
         List            =   "frm_student_batch.frx":001C
         TabIndex        =   3
         Text            =   "(Select)"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frm_student_batch.frx":0062
         Left            =   2160
         List            =   "frm_student_batch.frx":0072
         TabIndex        =   2
         Text            =   "(Select)"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         Left            =   2160
         TabIndex        =   0
         Text            =   "(Select Batch)"
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   "Select Season :"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Select Sem :"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Select Field :"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Selct Batch :"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm_student_batch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim a As String
Dim b As String
Dim count2 As Integer
Dim count3 As Integer
Dim count4 As Integer
Dim count5 As Integer
Private Sub Combo1_Click()
    count2 = 0
End Sub

Private Sub Combo2_Click()
    count2 = 0
    Combo3.Enabled = True
    If Combo2.Text = "OTHER" Then
        If count4 = 0 And Combo2.Text = "OTHER" Then
            count3 = 0
            Combo3.AddItem ("Sem-VII")
            Combo3.AddItem ("Sem-VIII")
        End If
    Else
        If count3 <> 1 Then
            count4 = 0
            count3 = 1
            Combo3.RemoveItem (7)
            Combo3.RemoveItem (6)
        End If
    End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Combo3_Click()
    count2 = 0
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command1_Click()
     Dim flag As Integer
            flag = 0
            List1.Clear
            List2.Clear
            Set rs1 = cnn.Execute("select sname,rollno from stud_sub_detail where sub1='" & Combo1.Text & "'or sub2='" & Combo1.Text & "'or sub3='" & Combo1.Text & "' or sub4='" & Combo1.Text & "'or sub5='" & Combo1.Text & "'or sub6='" & Combo1.Text & "'")
            Set rs2 = cnn.Execute("select sname,rollno from stud_sub_detail where season='" & Combo4.Text & "' and field='" & Combo2.Text & "' and sem='" & Combo3.Text & "'")
            If (Not rs1.EOF) Or (Not rs2.EOF) Then
                Do While Not rs1.EOF
                    Do While Not rs2.EOF
                        If rs1.Fields("sname") = rs2.Fields("sname") Then
                                List1.AddItem (rs1.Fields("sname"))
                                List2.AddItem (rs1.Fields("rollno"))
                                rs1.MoveNext
                                rs2.MoveNext
                                Exit Do
                        Else
                            rs1.MoveNext
                           ' rs2.MoveNext
                        End If
                    Loop
                    If rs2.EOF Then
                        Exit Do
                    End If
                Loop
            Else
                MsgBox "Students not found", vbExclamation + vbOKOnly, "Information"
            End If
            If List1.ListCount <> 0 Then
                Set rs = cnn.Execute("select * from stud_batch_details where batch='" & Combo1.Text & "'")
                For i = 0 To List1.ListCount - 1
                    If (Not rs.EOF) Then
                        Do While Not rs.EOF
                            If rs.Fields("sname") = List1.List(i) Then
                                flag = 1
                                Exit Do
                            End If
                            rs.MoveNext
                        Loop
                        If flag = 0 Then
                            List3.AddItem (List1.List(i))
                            If (Not rs.EOF) Then
                                rs.MoveNext
                            Else
                                rs.MoveFirst
                            End If
                        
                        Else
                            flag = 0
                            rs.MoveFirst
                        End If
                    Else
                        List3.AddItem (List1.List(i))
                    End If
                Next
            End If
            If List3.ListCount = 0 Then
                MsgBox "No Students found", vbOKOnly + vbExclamation, "Informationa"
            End If
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command2_Click()
    Dim count2 As Integer
    Dim count3 As Integer
    count3 = 0
    count2 = 0
    If count2 = 0 Then
        For i = 1 To List3.ListCount
            s = "insert into stud_batch_details (rollno,sname,batch,field,season) values ('" & List2.List(i - 1) & "','" & List3.List(i - 1) & "','" & Combo1.Text & "','" & Combo2.Text & "','" & Combo4.Text & "')"
            cnn.Execute s
            count2 = count2 + 1
        Next
        count3 = 1
    Else
        MsgBox "Students are already added", vbInformation + vbOKOnly, "information"
    End If
    If count2 = List3.ListCount Then
        MsgBox "Students added successfully", vbInformation + vbOKOnly, "Information"
        Unload Me
    End If
End Sub

Private Sub Form_Load()
     Me.Top = 2500
    Me.Left = 2900
    count2 = 0
    connect
    count3 = 0
    count4 = 1
    Set rs = cnn.Execute("select distinct  batch from batch_details order by batch")
    Do While Not rs.EOF
        Combo1.AddItem (rs.Fields("batch"))
        rs.MoveNext
    Loop
    a = Date
    b = Mid(a, 9, 10)
    Combo4.AddItem ("Winter " + b)
    Combo4.AddItem ("Summer " + b)
End Sub

