VERSION 5.00
Begin VB.Form frm_upadate_seat_no 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Seat Number :"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9585
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
   ScaleHeight     =   7335
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   4440
      TabIndex        =   13
      Top             =   0
      Width           =   4815
      Begin VB.ComboBox Combo4 
         Height          =   405
         Left            =   720
         TabIndex        =   18
         Text            =   "(Batch)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frm_upadate_seat_no.frx":0000
         Left            =   2640
         List            =   "frm_upadate_seat_no.frx":001C
         TabIndex        =   17
         Text            =   "(Sem)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frm_upadate_seat_no.frx":0062
         Left            =   120
         List            =   "frm_upadate_seat_no.frx":0072
         TabIndex        =   16
         Text            =   "(Field)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   4335
      Begin VB.OptionButton Option2 
         Caption         =   "By Batch"
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Semesters"
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student Information :"
      Height          =   6255
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   9495
      Begin VB.ComboBox Combo3 
         Height          =   405
         Left            =   3000
         TabIndex        =   19
         Text            =   "(Select Name)"
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Allocate"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1440
         TabIndex        =   2
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   2880
         TabIndex        =   1
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2880
         TabIndex        =   0
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   2880
         TabIndex        =   7
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   2880
         TabIndex        =   6
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   6480
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Enter Seat No :"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Student Name :"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Student RollNo :"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Field :"
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Select the Name :"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frm_upadate_seat_no"
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
        MsgBox "Please enter Name or rollno", vbQuestion + vbOKOnly, "Warning"
    Else
        Set rs = cnn.Execute("select * from admission where sname='" & Combo3.Text & "'")
        If (Not rs.EOF) Then
            Text2.Text = rs.Fields("sname")
            Text3.Text = rs.Fields("rollno")
            Text4.Text = rs.Fields("field")
            Command3.Enabled = True
        Else
            MsgBox "Student not found", vbOKOnly + vbInformation, "Information"
            Command3.Enabled = False
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
If Text5.Text = "" Then
    MsgBox "Please enter the seat no to allocate", vbOKOnly + vbCritical, "Warning"
Else
    s = "update admission set seatno='" & Text5.Text & "' where sname='" & Combo3.Text & "' and rollno='" & Text3.Text & "'"
    cnn.Execute s
    MsgBox "Seat no allocated successfully", vbOKOnly + vbInformation, "Successfull"
    Unload Me
End If
End Sub

Private Sub Form_Load()
     Me.Top = 1200
    Me.Left = 2500
    connect
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

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command3_Click
    End If
End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command3_Click
    End If
End Sub
