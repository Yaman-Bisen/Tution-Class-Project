VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_new_batch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create new batch :"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8220
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
   ScaleHeight     =   7005
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Batch-Details :"
      Height          =   6735
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   8055
      Begin VB.ComboBox Combo3 
         Height          =   405
         Left            =   2520
         TabIndex        =   5
         Text            =   "(Select)"
         Top             =   4320
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save Batch"
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Top             =   5520
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   2520
         TabIndex        =   4
         Top             =   3720
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   3000
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   96010241
         CurrentDate     =   43861
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   2280
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   96010241
         CurrentDate     =   43861
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frm_new_batch.frx":0000
         Left            =   2520
         List            =   "frm_new_batch.frx":000A
         TabIndex        =   1
         Text            =   "(Select)"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         Left            =   2520
         TabIndex        =   0
         Text            =   "(Select Batch)"
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "Season :"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Enter Timing :"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Ending Date :"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Starting Date :"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Week Days :"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Select Batch :"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frm_new_batch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim a As String

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command1_Click()
    If Combo1.Text = "(Select Batch)" Then
        MsgBox "Please select the batch", vbCritical + vbOKOnly, "Warning"
    ElseIf Combo2.Text = "(Select)" Then
        MsgBox "Please select the days", vbCritical + vbOKOnly, "warning"
    ElseIf Text1.Text = "" Then
        MsgBox "Please enter the timing of batch", vbCritical + vbOKOnly, "warning"
    ElseIf Combo3.Text = "(Select)" Then
         MsgBox "Please select the season", vbCritical + vbOKOnly, "warning"
    Else
        s = "insert into batch_details (batch,days,timing,starting_date,ending_date,season) values ('" & Combo1.Text & "','" & Combo2.Text & "','" & Text1.Text & "','" & DTPicker1.Value & "','" & DTPicker2.Value & "','" & Combo3.Text & "')"
        cnn.Execute s
        MsgBox "Batch created successfully", vbOKOnly + vbInformation, "Successfull"
        Unload Me
    End If
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    DTPicker2.Value = Date
     Me.Top = 1500
    Me.Left = 3700
    connect
    a = Date
    b = Mid(a, 9, 10)
    Combo3.AddItem ("Winter " + b)
    Combo3.AddItem ("Summer " + b)
    Set rs = cnn.Execute("select distinct  subject from subject_details order by subject")
    Do While Not rs.EOF
        Combo1.AddItem (rs.Fields("subject"))
        rs.MoveNext
    Loop
End Sub
