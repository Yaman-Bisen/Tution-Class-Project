VERSION 5.00
Begin VB.Form frm_sign_up 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sign Up :"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6435
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
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Enter User Details :"
      Height          =   5415
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   3000
         TabIndex        =   5
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Sign Up"
         Height          =   495
         Left            =   960
         TabIndex        =   4
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   3120
         TabIndex        =   0
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Return to Login....."
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Re-Ener Password :"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Enter Password :"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Username :"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frm_sign_up"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
If login_type = "user" Then
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
        MsgBox "Please enter whole information", vbInformation + vbOKOnly, "Warning"
    Else
        Set rs = cnn.Execute("select * from login where username='" & Text1.Text & "'")
        If (Not rs.EOF) Then
            MsgBox "Username already registered", vbExclamation + vbOKOnly, "Information"
        Else
            If Text2.Text = Text3.Text Then
                s = "insert into login values ('" & Text1.Text & "','" & Text2.Text & "','" & "Null" & "','" & "Null" & "')"
                cnn.Execute s
                MsgBox "Username registered successfully", vbInformation + vbOKOnly, "Information"
                Text1.Text = ""
                Text2.Text = ""
                Text3.Text = ""
            Else
                MsgBox "Password and Re-entered password does not match", vbCritical + vbOKOnly, "Warning"
            End If
        End If
    End If
ElseIf login_type = "admin" Then
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
        MsgBox "Please enter whole information", vbInformation + vbOKOnly, "Warning"
    Else
        Set rs = cnn.Execute("select * from login where username='" & Text1.Text & "'")
        If (Not rs.EOF) Then
            MsgBox "Username already registered", vbExclamation + vbOKOnly, "Information"
        Else
            If Text2.Text = Text3.Text Then
                s = "insert into login values ('" + "Null" + "','" + "Null" + "','" + Text1.Text + "','" + Text2.Text + "')"
                cnn.Execute s
                MsgBox "Username registered successfully", vbInformation + vbOKOnly, "Information"
                Text1.Text = ""
                Text2.Text = ""
                Text3.Text = ""
            Else
                MsgBox "Password and Re-entered password does not match", vbCritical + vbOKOnly, "Warning"
            End If
        End If
    End If
End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    connect
End Sub

Private Sub Label5_Click()
    frm_login.Show
    Unload Me
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

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub
