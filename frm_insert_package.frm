VERSION 5.00
Begin VB.Form frm_insert_package 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert New Package :"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10485
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
   ScaleHeight     =   7230
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Package Details :"
      Height          =   7215
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   10215
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   2760
         TabIndex        =   7
         ToolTipText     =   "if not enter '0'"
         Top             =   5280
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   495
         Left            =   1320
         TabIndex        =   8
         Top             =   6120
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   2760
         TabIndex        =   6
         Top             =   4680
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   2760
         TabIndex        =   5
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2760
         TabIndex        =   4
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   375
         Left            =   5520
         TabIndex        =   15
         Top             =   2520
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   4335
         Left            =   6600
         TabIndex        =   14
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   2760
         TabIndex        =   3
         Top             =   2520
         Width           =   2655
      End
      Begin VB.ComboBox Combo3 
         Height          =   405
         ItemData        =   "frm_insert_package.frx":0000
         Left            =   2760
         List            =   "frm_insert_package.frx":000A
         TabIndex        =   2
         Text            =   "(Select)"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frm_insert_package.frx":002A
         Left            =   2760
         List            =   "frm_insert_package.frx":003A
         TabIndex        =   0
         Text            =   "(Select Field)"
         Top             =   600
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   405
         ItemData        =   "frm_insert_package.frx":005A
         Left            =   2760
         List            =   "frm_insert_package.frx":0076
         TabIndex        =   1
         Text            =   "(Select Sem)"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Package Fees :"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "3rd Intallment :"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "2nd Installment :"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "1st Installment :"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Enter Subject :"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Select Package :"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Select Field :"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Select Sem :"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_insert_package"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim count1 As Integer
Dim count2 As Integer

Private Sub Command1_Click()
    List1.AddItem (Text1.Text)
    Text1.Text = ""
End Sub

Private Sub Command2_Click()
    If Combo1.Text = "(Select Field)" Then
        MsgBox "Please select the field", vbExclamation + vbOKOnly, "Warning"
    ElseIf Combo2.Text = "(Select Sem)" Then
        MsgBox "Please select semester", vbExclamation + vbOKOnly, "Warning"
    ElseIf Combo3.Text = "(Select)" Then
        MsgBox "Please select package type", vbExclamation + vbOKOnly, "Warning"
    ElseIf List1.ListCount = 0 Then
        MsgBox "Please add the subjects", vbExclamation + vbOKOnly, "Warning"
    ElseIf Text2.Text = "" Then
        MsgBox "Please enter 1st installment", vbExclamation + vbOKOnly, "Warning"
    ElseIf Text3.Text = "" Then
        MsgBox "Please enter 2nd installment", vbExclamation + vbOKOnly, "Warning"
    ElseIf Text4.Text = "" Then
        MsgBox "Please enter 3rd installment", vbExclamation + vbOKOnly, "Warning"
    ElseIf Text5.Text = "" Then
        MsgBox "Please enter package fee", vbExclamation + vbOKOnly, "Warning"
    Else
        s = "insert into package_details (field,sem,package_type,price,first_installment,second_installment,third_installment) values ('" & Combo1.Text & "','" & Combo2.Text & "','" & Combo3.Text & "','" & Text5.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "')"
        cnn.Execute s
        MsgBox "Package saved successfully", vbInformation + vbOKOnly, "Successfull"
        If List1.ListCount = 1 Then
            s = "update package_details set subject1='" & List1.List(0) & "',subject2='" & "Null" & "',subject3='" & "Null" & "',subject4='" & "Null" & "',subject5='" & "Null" & "',subject6='" & "Null" & "' where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "' and Package_type='" & Combo3.Text & "'"
            cnn.Execute s
        ElseIf List1.ListCount = 2 Then
            s = "update package_details set subject1='" & List1.List(0) & "',subject2='" & List1.List(1) & "',subject3='" & "Null" & "',subject4='" & "Null" & "',subject5='" & "Null" & "',subject6='" & "Null" & "' where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "' and Package_type='" & Combo3.Text & "'"
            cnn.Execute s
        ElseIf List1.ListCount = 3 Then
            s = "update package_details set subject1='" & List1.List(0) & "',subject2='" & List1.List(1) & "',subject3='" & List1.List(2) & "',subject4='" & "Null" & "',subject5='" & "Null" & "',subject6='" & "Null" & "' where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "' and Package_type='" & Combo3.Text & "'"
            cnn.Execute s
        ElseIf List1.ListCount = 4 Then
            s = "update package_details set subject1='" & List1.List(0) & "',subject2='" & List1.List(1) & "',subject3='" & List1.List(2) & "',subject4='" & List1.List(3) & "',subject5='" & "Null" & "',subject6='" & "Null" & "' where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "' and Package_type='" & Combo3.Text & "'"
            cnn.Execute s
        ElseIf List1.ListCount = 5 Then
            s = "update package_details set subject1='" & List1.List(0) & "',subject2='" & List1.List(1) & "',subject3='" & List1.List(2) & "',subject4='" & List1.List(3) & "',subject5='" & List1.List(4) & "',subject6='" & "Null" & "' where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "' and Package_type='" & Combo3.Text & "'"
            cnn.Execute s
        ElseIf List1.ListCount = 6 Then
            s = "update package_details set subject1='" & List1.List(0) & "',subject2='" & List1.List(1) & "',subject3='" & List1.List(2) & "',subject4='" & List1.List(3) & "',subject5='" & List1.List(4) & "',subject6='" & List1.List(5) & "' where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "' and Package_type='" & Combo3.Text & "'"
            cnn.Execute s
        End If
        Unload Me
    End If
    End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command2_Click
    End If
End Sub

Private Sub Form_Load()
     Me.Top = 1500
    Me.Left = 3000
    count1 = 0
    count2 = 1
    connect
End Sub
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

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command2_Click
    End If
End Sub
