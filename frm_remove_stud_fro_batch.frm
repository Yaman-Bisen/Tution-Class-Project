VERSION 5.00
Begin VB.Form frm_remove_stud_fro_batch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remove Student From Batch :"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5940
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
   ScaleHeight     =   5715
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Student Details :"
      Height          =   3975
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   5895
      Begin VB.CommandButton Command1 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1920
         TabIndex        =   12
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         Left            =   2160
         TabIndex        =   7
         Text            =   "(Select)"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label7 
         Height          =   495
         Left            =   2160
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Field :"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label5 
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Roll Number :"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Select Student :"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.ComboBox Combo4 
         Height          =   405
         Left            =   2160
         TabIndex        =   4
         Text            =   "(Select)"
         Top             =   960
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         Left            =   2160
         TabIndex        =   2
         Text            =   "(Select Batch)"
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Select Season :"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Select Batch :"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm_remove_stud_fro_batch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset



Private Sub Combo2_Click()
    Set rs = cnn.Execute("select * from stud_batch_details where batch='" & Combo1.Text & "'and season='" & Combo4.Text & "' and sname='" & Combo2.Text & "'")
    If (Not rs.EOF) Then
        Label5.Caption = rs.Fields("rollno")
        Label7.Caption = rs.Fields("field")
        Command1.Enabled = True
    Else
        MsgBox "Student information is not available", vbOKOnly + vbInformation, "Information"
    End If
End Sub

Private Sub Combo4_Click()
    If Combo1.Text = "(Select Batch)" Then
        MsgBox "Please select batch first", vbExclamation + vbOKOnly, "Warning"
    Else
        Set rs = cnn.Execute("select * from stud_batch_details where batch='" & Combo1.Text & "'and season='" & Combo4.Text & "'")
        If (Not rs.EOF) Then
            Do While Not rs.EOF
                Combo2.AddItem (rs.Fields("sname"))
                rs.MoveNext
            Loop
        Else
            MsgBox "Students not found for selected batch", vbOKOnly + vbInformation, "Information"
        End If
    End If
End Sub

Private Sub Command1_Click()
    Dim res As VbMsgBoxResult
    res = MsgBox("Are you sure to delete sudent from batch", vbOKOnly + vbQuestion, "Asking")
    If res = vbYes Then
        s = "Delete from stud_batch_details where rollno='" & Label5.Caption & "'and sname='" & Combo2.Text & "'and batch='" & Combo1.Text & "'and season='" & Combo4.Text & "'and field='" & Label7.Caption & "'"
        cnn.Execute s
        MsgBox "Student removed from batch successfully", vbInformation + vbOKOnly, "Information"
        Unload Me
    Else
        MsgBox "You Cancelled the operation", vbOKOnly + vbInformation, "Information"
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Left = 4500
    Me.Top = 1500
    connect
    Command1.Enabled = True
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

