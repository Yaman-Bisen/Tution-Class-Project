VERSION 5.00
Begin VB.Form frm_installment_bal_rpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pending Installments :"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8490
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
   ScaleHeight     =   5580
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Select Details :"
      Height          =   5535
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8415
      Begin VB.ComboBox Combo3 
         Height          =   405
         ItemData        =   "frm_installment_bal_rpt.frx":0000
         Left            =   3120
         List            =   "frm_installment_bal_rpt.frx":0010
         TabIndex        =   1
         Text            =   "(Select)"
         Top             =   1920
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Search"
         Height          =   495
         Left            =   1920
         TabIndex        =   3
         Top             =   3600
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frm_installment_bal_rpt.frx":0030
         Left            =   3120
         List            =   "frm_installment_bal_rpt.frx":003D
         TabIndex        =   2
         Text            =   "(Select)"
         Top             =   2760
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         Left            =   3120
         TabIndex        =   0
         Text            =   "(Select)"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Select Field :"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Select Installment :"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Select Batch :"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frm_installment_bal_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim a As String


Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
        DataEnvironment1.rsCommand4.Close
        DataEnvironment1.rsCommand3.Close
        DataEnvironment1.rsCommand5.Close
    On Error GoTo 0
        Set rs = cnn.Execute("select * from package_details where field='" & Combo3.Text & "' and subject1='" & Combo1.Text & "' or subject2='" & Combo1.Text & "' or subject3='" & Combo1.Text & "' or subject4='" & Combo1.Text & "' or subject5='" & Combo1.Text & "' or subject6='" & Combo1.Text & "'")
            Do While Not rs.EOF
                If rs.Fields("field") = Combo3.Text Then
                    If Combo2.Text = "1st Installment" Then
                        a = rs.Fields("first_installment")
                        DataEnvironment1.Command4 a
                        rpt_1st_balance.Show
                        Exit Do
                    ElseIf Combo2.Text = "2nd Installment" Then
                        a = rs.Fields("second_installment")
                        DataEnvironment1.Command3 a
                        rpt_balance_installment.Show
                        Exit Do
                    ElseIf Combo2.Text = "3rd Installment" Then
                        a = rs.Fields("third_installment")
                        DataEnvironment1.Command5 a
                        rpt_3rd_ins_balance.Show
                        Exit Do
                    End If
                End If
                rs.MoveNext
            Loop
            Unload Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Form_Load()
     Me.Top = 2000
    Me.Left = 3200
    connect
    Set rs = cnn.Execute("select distinct  batch from batch_details order by batch")
    Do While Not rs.EOF
        Combo1.AddItem (rs.Fields("batch"))
        rs.MoveNext
    Loop
End Sub
