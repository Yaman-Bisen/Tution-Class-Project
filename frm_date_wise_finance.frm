VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_date_wise_finance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Date wise finance information :"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6180
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
   ScaleHeight     =   2940
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton Command1 
         Caption         =   "&View"
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   1680
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2880
         TabIndex        =   0
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   96010241
         CurrentDate     =   43869
      End
      Begin VB.Label Label1 
         Caption         =   "Select Date :"
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_date_wise_finance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim b As String
Private Sub Command1_Click()
    On Error Resume Next
        DataEnvironment1.rsCommand12.Close
    On Error GoTo 0
    Set rs = cnn.Execute("select sum(price) as total from bill_view where dates ='" & DTPicker1.Value & "'")
    If Not rs.EOF Then
        If (Not rs![total]) Then
            b = rs![total]
            a = DTPicker1.Value
            DataEnvironment1.Command12 b, a
            rpt_date_wise_finance.Show
            Unload Me
        Else
            MsgBox "No data available for selected date", vbInformation + vbOKOnly, "Information"
        End If
    End If
    
    
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    connect
     Me.Top = 2500
    Me.Left = 3500
End Sub
