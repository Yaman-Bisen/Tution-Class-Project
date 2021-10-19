VERSION 5.00
Begin VB.Form frm_admitted_stud_report 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Admitted Students :"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5430
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
   ScaleHeight     =   3840
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Enter Details :"
      Height          =   3735
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "&Search"
         Height          =   495
         Left            =   1680
         TabIndex        =   2
         Top             =   2520
         Width           =   1815
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frm_admitted_stud_report.frx":0000
         Left            =   2040
         List            =   "frm_admitted_stud_report.frx":001C
         TabIndex        =   1
         Text            =   "(Select Sem)"
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frm_admitted_stud_report.frx":0062
         Left            =   2040
         List            =   "frm_admitted_stud_report.frx":0072
         TabIndex        =   0
         Text            =   "(Select Field)"
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Select Sem :"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
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
Attribute VB_Name = "frm_admitted_stud_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command1_Click()
 '   On Error Resume Next
   '  If DataEnvironment1.rsCommand1.State = 1 Then DataEnvironment1.rsCommand1.Close
  '  On Error GoTo 0
    n = Combo1.Text
    m = Combo2.Text
    DataEnvironment1.Command1 n, m
    rpt_admitted_stud.Show
    'DataEnvironment1.rsCommand1.Close
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
    count1 = 0
    count2 = 0
    Combo2.Enabled = False
    'If DataEnvironment1.rsCommand1.State = 1 Then DataEnvironment1.rsCommand1.Close
    If DataEnvironment1.Connection1.State = 1 Then DataEnvironment1.Connection1.Close
    DataEnvironment1.Connection1.Open
End Sub
