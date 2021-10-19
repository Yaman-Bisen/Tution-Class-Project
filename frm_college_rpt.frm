VERSION 5.00
Begin VB.Form frm_college_rpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selected college students :"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5805
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
   ScaleHeight     =   4230
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Select Details :"
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "&Search"
         Height          =   495
         Left            =   1680
         TabIndex        =   7
         Top             =   3120
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   405
         Left            =   2640
         TabIndex        =   2
         Text            =   "(Select)"
         Top             =   2040
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frm_college_rpt.frx":0000
         Left            =   2640
         List            =   "frm_college_rpt.frx":001C
         TabIndex        =   1
         Text            =   "(Select)"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frm_college_rpt.frx":0062
         Left            =   2640
         List            =   "frm_college_rpt.frx":0072
         TabIndex        =   0
         Text            =   "(Select)"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Select College :"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Select Sem :"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Select Field :"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm_college_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim count1 As Integer
Dim count2 As Integer

Private Sub Combo1_Click()
    Combo2.Enabled = True
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
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
        DataEnvironment1.rsCommand2.Close
    On Error GoTo 0
    n = Combo1.Text
    m = Combo2.Text
    o = Combo3.Text
    DataEnvironment1.Command2 n, m, o
    rpt_college_wise.Show
    Unload Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Form_Load()
     Me.Top = 2500
    Me.Left = 4000
    connect
    count1 = 0
    count2 = 1
    Set rs = cnn.Execute("select * from college'")
    Do While Not rs.EOF
        Combo3.AddItem (rs.Fields("cname"))
        rs.MoveNext
    Loop
    Combo2.Enabled = False
End Sub
