VERSION 5.00
Begin VB.Form frm_batch_wise 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batch Wise Report :"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6765
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
   ScaleHeight     =   3990
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.ComboBox Combo1 
         Height          =   405
         Left            =   2400
         TabIndex        =   3
         Text            =   "(Select Batch)"
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&View"
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   2520
         Width           =   1815
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         Left            =   2400
         TabIndex        =   1
         Text            =   "(Select)"
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Select Batch :"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Select Season :"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frm_batch_wise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
        DataEnvironment1.rsCommand7.Close
    On Error GoTo 0
   m = Combo1.Text
   n = Combo2.Text
   DataEnvironment1.Command7 m, n
   rpt_batch_wise.Show
   Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = 2500
    Me.Left = 3500
    a = Date
    b = Mid(a, 9, 10)
    Combo2.AddItem ("Winter " + b)
    Combo2.AddItem ("Summer " + b)
    connect
    Set rs = cnn.Execute("select distinct  batch from batch_details order by batch")
    If (Not rs.EOF) Then
        Do While Not rs.EOF
            Combo1.AddItem (rs.Fields("batch"))
            rs.MoveNext
        Loop
    Else
        MsgBox "Batches are not created yet", vbExclamation + vbOKOnly, "Information"
    End If
End Sub
