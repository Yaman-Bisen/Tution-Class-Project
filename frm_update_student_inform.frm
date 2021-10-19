VERSION 5.00
Begin VB.Form frm_update_student_inform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Student Information :"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9555
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
   ScaleHeight     =   8205
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   4440
      TabIndex        =   19
      Top             =   0
      Width           =   5055
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frm_update_student_inform.frx":0000
         Left            =   120
         List            =   "frm_update_student_inform.frx":0010
         TabIndex        =   24
         Text            =   "(Field)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frm_update_student_inform.frx":0030
         Left            =   2760
         List            =   "frm_update_student_inform.frx":004C
         TabIndex        =   23
         Text            =   "(Sem)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox Combo4 
         Height          =   405
         Left            =   840
         TabIndex        =   22
         Text            =   "(Batch)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   4095
      Begin VB.OptionButton Option1 
         Caption         =   "By Semesters"
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "By Batch"
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Update Student Information :"
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   9495
      Begin VB.ComboBox Combo3 
         Height          =   405
         Left            =   3000
         TabIndex        =   25
         Text            =   "(Select Name)"
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Update"
         Height          =   495
         Left            =   1680
         TabIndex        =   17
         Top             =   6360
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         Height          =   405
         Left            =   2760
         TabIndex        =   16
         Top             =   5520
         Width           =   3375
      End
      Begin VB.TextBox Text15 
         Height          =   405
         Left            =   2760
         TabIndex        =   15
         Top             =   4800
         Width           =   3375
      End
      Begin VB.TextBox Text14 
         Height          =   405
         Left            =   2760
         TabIndex        =   14
         Top             =   4080
         Width           =   3375
      End
      Begin VB.TextBox Text13 
         Height          =   405
         Left            =   2760
         TabIndex        =   13
         Top             =   3480
         Width           =   3375
      End
      Begin VB.TextBox Text12 
         Height          =   405
         Left            =   2760
         TabIndex        =   12
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox Text11 
         Enabled         =   0   'False
         Height          =   405
         Left            =   2760
         TabIndex        =   11
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox Text10 
         Height          =   405
         Left            =   2760
         TabIndex        =   10
         Top             =   1440
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   405
         Left            =   6480
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Pin-code :"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Landmark :"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Area :"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Parent Mob No :"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Personal No :"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Roll No :"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Student Name :"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Select Student Name :"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frm_update_student_inform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim count1 As Integer
Dim count2 As Integer

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
        MsgBox "Please select the name of the student", vbCritical + vbOKOnly, "Warning"
    Else
        Set rs = cnn.Execute("select * from admission where sname='" & Combo3.Text & "'")
        If (Not rs.EOF) Then
            Text10.Text = rs.Fields("sname")
            Text11.Text = rs.Fields("rollno")
            Text12.Text = rs.Fields("selfmobno")
            Text13.Text = rs.Fields("parentmobno")
            Text14.Text = rs.Fields("area")
            Text15.Text = rs.Fields("landmark")
            Text16.Text = rs.Fields("pincode")
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

Private Sub Command3_Click()
If Len(Text12.Text) < 10 Then
    MsgBox "Mobile numbers should be of 10 digits", vbCritical + vbOKOnly, "Warning"
Else
    If Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Or Text16.Text = "" Then
        MsgBox "Please enter all the information", vbCritical + vbOKOnly, "Warning"
    Else
        s = "update admission set sname='" & Text10.Text & "',selfmobno='" & Text12.Text & "',parentmobno='" & Text13.Text & "',area='" & Text14.Text & "',landmark='" & Text15.Text & "',pincode='" & Text16.Text & "' where rollno='" & Text11.Text & "'"
        cnn.Execute s
        MsgBox "Information updated successfully", vbInformation + vbOKOnly, "Information"
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
     Me.Top = 700
    count1 = 0
    count2 = 0
    Me.Left = 2500
   connect
    Set rs = cnn.Execute("select * from batch_details")
    If (Not rs.EOF) Then
        Do While Not rs.EOF
            Combo4.AddItem (rs.Fields("batch"))
            rs.MoveNext
        Loop
    End If
End Sub



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

Private Sub Text1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Call Command1_Click
'    End If
End Sub


Private Sub Text15_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command3_Click
    End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Call Command1_Click
'    End If
End Sub
