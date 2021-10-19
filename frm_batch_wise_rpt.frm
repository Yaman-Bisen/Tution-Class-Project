VERSION 5.00
Begin VB.Form frm_batch_wise_rpt 
   Caption         =   "Batch Wise Students :"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7020
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
   MDIChild        =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   7020
   Begin VB.Frame Frame1 
      Caption         =   "Select Batch Details :"
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   495
         Left            =   1680
         TabIndex        =   5
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   2400
         TabIndex        =   4
         ToolTipText     =   "Winter 2020"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         Left            =   2400
         TabIndex        =   2
         Text            =   "(Select Batch)"
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Enter Season :"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Select Batch :"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frm_batch_wise_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
    m = Combo1.Text
    n = Text1.Text
    DataEnvironment1.Command7 m, n
    rpt_batch_wise.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Move (MDIForm1.ScaleWidth - Me.Width)
    connect
    Set rs = cnn.Execute("select distinct  batch from batch_details order by batch")
    Do While Not rs.EOF
        Combo1.AddItem (rs.Fields("batch"))
        rs.MoveNext
    Loop
End Sub
