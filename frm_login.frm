VERSION 5.00
Begin VB.Form frm_login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login to D'Soft :"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8025
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   4575
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8055
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   3720
         TabIndex        =   0
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   4800
         TabIndex        =   4
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Login"
         Height          =   495
         Left            =   3120
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   3720
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "  :-)"
         Height          =   375
         Left            =   6240
         TabIndex        =   10
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sign Up ....."
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   6360
         TabIndex        =   3
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password :"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Username :"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs As New ADODB.Recordset

Private Sub Command1_Click()
    If login_type = "user" Then
        rs.Open "select * from login where username = '" + Text1.Text + "' and password = '" + Text2.Text + "'", cnn
        If (Not rs.EOF) Then
            MDIForm1.Show
            Unload Me
        Else
            MsgBox "Invalid user", vbCritical + vbOKOnly, "Warning"
        End If
        rs.Close
    ElseIf login_type = "admin" Then
        rs.Open "select * from login where adminuser = '" + Text1.Text + "' and adminpass = '" + Text2.Text + "'", cnn
        If (Not rs.EOF) Then
            MDIForm1.Show
            Unload Me
        Else
            MsgBox "Invalid user", vbCritical + vbOKOnly, "Warning"
        End If
        rs.Close
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    connect
    If login_type = "admin" Then
        Label4.Caption = "Login As Admin..."
        Label5.Caption = "Switch To..."
    ElseIf login_type = "user" Then
        Label4.Caption = "Login As User..."
        Label5.Caption = "Switch To..."
    End If
End Sub
Private Sub Label3_Click()
   sign = "yes"
   Form3.Show
   Unload Me
End Sub

Private Sub Label5_Click()
    Form3.Show
    Unload Me
End Sub

Private Sub Label6_Click()
    If Text2.PasswordChar = "*" Then
        Text2.PasswordChar = ""
    Else
        Text2.PasswordChar = "*"
    End If
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
