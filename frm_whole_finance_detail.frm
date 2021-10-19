VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Total amount of admitted students :"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6270
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
   ScaleHeight     =   3630
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_whole_finance_detail 
      Caption         =   "Select Dates :"
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "&Search"
         Height          =   495
         Left            =   1680
         TabIndex        =   3
         Top             =   2400
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   96010241
         CurrentDate     =   43864
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   96010241
         CurrentDate     =   43864
      End
      Begin VB.Label Label2 
         Caption         =   "Select to Date :"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Select From Date :"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
    On Error Resume Next
        DataEnvironment1.rsCommand6.Close
    On Error GoTo 0
    n = DTPicker1.Value
    m = DTPicker2.Value
    DataEnvironment1.Command6 n, m, n, m
    'rpt_whole_finance.Show
    rpt_finance_whole.Show
    Unload Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub DTPicker2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    Me.Top = 2500
    Me.Left = 3800
    connect
End Sub
