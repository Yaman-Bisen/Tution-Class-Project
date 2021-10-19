VERSION 5.00
Begin VB.Form frm_search_student 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Student Details :"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10035
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
   ScaleHeight     =   8610
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   4440
      TabIndex        =   26
      Top             =   0
      Width           =   4575
      Begin VB.ComboBox Combo4 
         Height          =   405
         Left            =   960
         TabIndex        =   30
         Text            =   "(Batch)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frm_search_student.frx":0000
         Left            =   2400
         List            =   "frm_search_student.frx":001C
         TabIndex        =   28
         Text            =   "(Sem)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frm_search_student.frx":0062
         Left            =   240
         List            =   "frm_search_student.frx":0072
         TabIndex        =   27
         Text            =   "(Field)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   4335
      Begin VB.OptionButton Option2 
         Caption         =   "By Batch"
         Height          =   285
         Left            =   2640
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Semesters"
         Height          =   285
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student Details :"
      Height          =   7695
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   9975
      Begin VB.ComboBox Combo3 
         Height          =   405
         Left            =   3000
         TabIndex        =   29
         Text            =   "(Select Name)"
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   8280
         TabIndex        =   0
         Top             =   7200
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   2760
         TabIndex        =   10
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   2760
         TabIndex        =   8
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   2760
         TabIndex        =   6
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2760
         TabIndex        =   4
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label17 
         Height          =   375
         Left            =   2760
         TabIndex        =   22
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Balance Fees :"
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label15 
         Height          =   375
         Left            =   2760
         TabIndex        =   20
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Total Fees :"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label13 
         Height          =   375
         Left            =   2760
         TabIndex        =   18
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label Label12 
         Caption         =   "Package Type :"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Label11 
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Top             =   7200
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "3rd Installment :"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   7200
         Width           =   1935
      End
      Begin VB.Label Label9 
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   6600
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "2nd Installment :"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   6600
         Width           =   2055
      End
      Begin VB.Label Label7 
         Height          =   375
         Left            =   2760
         TabIndex        =   12
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "1st Installment :"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   6000
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "College :"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Field :"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Student RollNo :"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Student Name :"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Select Student Name :"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frm_search_student"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim count1 As Integer
Dim count2 As Integer
Dim a As Integer
Dim b As Integer
Dim c As Integer

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
    If Combo2.Text <> "(Sem)" Then
        Set rs = cnn.Execute("select * from admission where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "'")
        Do While Not rs.EOF
            Combo3.AddItem (rs.Fields("sname"))
            rs.MoveNext
        Loop
    End If
End Sub

Private Sub Combo2_Click()
    Combo3.Clear
    If Combo1.Text = "(Field)" Then
        MsgBox "Please select the Field", vbQuestion + vbOKOnly, "Warning"
    ElseIf Combo2.Text = "(Sem)" Then
        MsgBox "Please select the Semester", vbQuestion + vbOKOnly, "Warning"
    Else
        Set rs = cnn.Execute("select * from admission where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "'")
        Do While Not rs.EOF
            Combo3.AddItem (rs.Fields("sname"))
            rs.MoveNext
        Loop
    End If
End Sub



Private Sub Combo3_Click()
    If Combo3.Text = "(Select Name)" Then
        MsgBox "Please select the name of student", vbQuestion + vbOKOnly, "Warning"
    Else
        Set rs = cnn.Execute("select * from admission where sname='" & Combo3.Text & "'")
        If (Not rs.EOF) Then
            Text2.Text = rs.Fields("sname")
            Text3.Text = rs.Fields("rollno")
            Text4.Text = rs.Fields("field")
            Text5.Text = rs.Fields("college")
            Label13.Caption = rs.Fields("subject")
            Label15.Caption = rs.Fields("totalfee")
            If Label13.Caption = "Select Subjects" Then
                Label10.Visible = False
                Label11.Visible = False
            Else
                Label10.Visible = True
                Label11.Visible = True
            End If
            Set rs1 = cnn.Execute("select * from trans where sname='" & Text2.Text & "'")
            If (Not rs1.EOF) Then
                Label17.Caption = rs1.Fields("balance")
                a = Val(rs1.Fields("firstinstallment"))
                b = Val(rs1.Fields("secondinstallment"))
                c = Val(rs1.Fields("thirdinstallment"))
                If a = 0 Then
                    Label7.Caption = "Not Paid"
                ElseIf a > 0 Then
                    Label7.Caption = a
                End If
                If b = 0 Then
                    Label9.Caption = "Not Paid"
                ElseIf b > 0 Then
                    Label9.Caption = b
                End If
                If c = 0 Then
                    Label11.Caption = "Not Paid"
                ElseIf c > 0 Then
                    Label11.Caption = c
                End If
            End If
        End If
    End If
End Sub

Private Sub Combo4_Click()
    Combo3.Clear
    Set rs = cnn.Execute("Select * from stud_batch_details where batch='" & Combo4.Text & "'")
    Do While Not rs.EOF
        Combo3.AddItem (rs.Fields("sname"))
        rs.MoveNext
    Loop
End Sub


Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Option1.Value = False
    Option2.Value = False
     Me.Top = 600
    Me.Left = 2500
    count1 = 0
    count2 = 1
    connect
    Set rs = cnn.Execute("select * from batch_details")
    If (Not rs.EOF) Then
        Do While Not rs.EOF
            Combo4.AddItem (rs.Fields("batch"))
            rs.MoveNext
        Loop
    End If
End Sub

'Private Sub Text1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Call Command1_Click
'    End If
'End Sub

Private Sub Option1_Click()
    Combo1.Visible = True
    Combo2.Visible = True
    Combo4.Visible = False
End Sub

Private Sub Option2_Click()
    Combo1.Visible = False
    Combo2.Visible = False
    Combo4.Visible = True
End Sub
