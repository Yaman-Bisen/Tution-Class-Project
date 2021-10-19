VERSION 5.00
Begin VB.Form frm_batch_student 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "College Students In Batch :"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7050
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
   ScaleHeight     =   4020
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6975
      Begin VB.ComboBox Combo3 
         Height          =   405
         Left            =   2760
         TabIndex        =   2
         Text            =   "(Select)"
         Top             =   1920
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&View"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frm_batch_student.frx":0000
         Left            =   2760
         List            =   "frm_batch_student.frx":0002
         TabIndex        =   1
         Text            =   "(Select College)"
         Top             =   1200
         Width           =   3615
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         Left            =   2760
         TabIndex        =   0
         Text            =   "(Select Batch)"
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "Select Season :"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Select College :"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Select Batch :"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frm_batch_student"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
        DataEnvironment1.rsCommand9.Close
    On Error GoTo 0
    m = Combo1.Text
    n = Combo2.Text
    o = Combo3.Text
    DataEnvironment1.Command9 m, n, o
    rpt_batch_student.Show
    Unload Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Form_Load()
     Me.Top = 2500
    Me.Left = 3800
    Me.Refresh
    connect
    Set rs = cnn.Execute("select distinct  batch from batch_details order by batch")
    Do While Not rs.EOF
        Combo1.AddItem (rs.Fields("batch"))
        rs.MoveNext
    Loop
    Set rs = cnn.Execute("select * from college'")
    Do While Not rs.EOF
        Combo2.AddItem (rs.Fields("cname"))
        rs.MoveNext
    Loop
    a = Date
    b = Mid(a, 9, 10)
    Combo3.AddItem ("Winter " + b)
    Combo3.AddItem ("Summer " + b)
End Sub
