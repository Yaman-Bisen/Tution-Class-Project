VERSION 5.00
Begin VB.Form frm_insert_subject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Subject :"
   ClientHeight    =   4410
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Subject Details :"
      Height          =   4335
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   495
         Left            =   1680
         TabIndex        =   4
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2760
         TabIndex        =   3
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   2760
         TabIndex        =   2
         Top             =   1770
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   405
         ItemData        =   "frm_insert_subject.frx":0000
         Left            =   2760
         List            =   "frm_insert_subject.frx":001C
         TabIndex        =   1
         Text            =   "(Select Sem)"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frm_insert_subject.frx":0062
         Left            =   2760
         List            =   "frm_insert_subject.frx":0072
         TabIndex        =   0
         Text            =   "(Select Field)"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Enter Price :"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Subject Name :"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Select Sem :"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Select Field :"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm_insert_subject"
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

Private Sub Command1_Click()
    If Combo1.Text = "(Select Field)" Then
        MsgBox "Please select the field", vbExclamation + vbOKOnly, "Warning"
    ElseIf Combo2.Text = "(Select Sem)" Then
        MsgBox "Please select semester", vbExclamation + vbOKOnly, "Warning"
    ElseIf Text1.Text = "" Then
        MsgBox "Please enter the name of subject", vbExclamation + vbOKOnly, "Warning"
    ElseIf Text2.Text = "" Then
        MsgBox "Please enter price for subject", vbExclamation + vbOKOnly, "Warning"
    Else
        s = "insert into subject_details (years,sem,subject,cost) values ('" & Combo1.Text & "','" & Combo2.Text & "','" & Text1.Text & "','" & Text2.Text & "')"
        cnn.Execute s
        MsgBox "Subject Added Successfully", vbInformation + vbOKOnly, "Successfull"
        Unload Me
    End If
End Sub

Private Sub Form_Load()
     Me.Top = 2500
    Me.Left = 4000
    count1 = 0
    count2 = 1
    connect
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub
