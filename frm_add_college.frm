VERSION 5.00
Begin VB.Form frm_add_college 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add College :"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5865
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
   ScaleHeight     =   2265
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   3000
         TabIndex        =   0
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Enter College Name :"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frm_add_college"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
    If Text1.Text = "" Then
        MsgBox "Please enter the college name", vbCritical + vbOKOnly, "Warning"
    Else
        s = "insert into college values ('" & Text1.Text & "')"
        cnn.Execute s
        MsgBox "College added successfully", vbInformation + vbOKOnly, "Information"
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 3000
    Me.Left = 3800
    connect
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub
