VERSION 5.00
Begin VB.Form frm_batch_view 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batch Information :"
   ClientHeight    =   3090
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "View Batches :"
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5775
      Begin VB.ComboBox Combo7 
         Height          =   405
         Left            =   2520
         TabIndex        =   3
         Text            =   "(Select)"
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&View"
         Height          =   495
         Left            =   1800
         TabIndex        =   0
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Select Season :"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frm_batch_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
        DataEnvironment1.rsCommand10.Close
    On Error GoTo 0
If Combo7.Text = "(Select)" Then
    MsgBox "Please select the season", vbExclamation + vbOKOnly, "Warning"
Else
    n = Combo7.Text
    DataEnvironment1.Command10 n
    rpt_view_batch.Show
    Unload Me
End If
End Sub

Private Sub Form_Load()
     Me.Top = 2500
    Me.Left = 4000
    connect
    a = Date
    b = Mid(a, 9, 10)
    Combo7.AddItem ("Winter " + b)
    Combo7.AddItem ("Summer " + b)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub
