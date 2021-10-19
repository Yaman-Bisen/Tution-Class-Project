VERSION 5.00
Begin VB.Form frm_package_details 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Packages"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6795
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
   ScaleHeight     =   2790
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Package details :"
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command1 
         Caption         =   "&View"
         Height          =   495
         Left            =   2160
         TabIndex        =   2
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frm_package_details.frx":0000
         Left            =   3360
         List            =   "frm_package_details.frx":000A
         TabIndex        =   0
         Text            =   "(Select)"
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Select Package Type :"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frm_package_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
        DataEnvironment1.rsCommand11.Close
    On Error GoTo 0
If Combo1.Text = "(Select)" Then
    MsgBox "Please selet the package type", vbExclamation + vbOKOnly, "Warning"
Else
    m = Combo1.Text
    DataEnvironment1.Command11 m
    rpt_view_packages.Show
    Unload Me
End If
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Form_Load()
     Me.Top = 2500
    Me.Left = 3000
End Sub
