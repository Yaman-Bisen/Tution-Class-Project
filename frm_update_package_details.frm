VERSION 5.00
Begin VB.Form frm_update_package_details 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Package Details :"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9435
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
   ScaleHeight     =   7365
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Package Details :"
      Height          =   7335
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton Command3 
         Caption         =   "&Update"
         Height          =   495
         Left            =   2880
         TabIndex        =   20
         Top             =   6240
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         Height          =   405
         Left            =   2400
         TabIndex        =   5
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         Height          =   405
         Left            =   2400
         TabIndex        =   6
         Top             =   4560
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   405
         Left            =   2400
         TabIndex        =   7
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   1560
         TabIndex        =   4
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6720
         TabIndex        =   19
         Top             =   5040
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   2280
         TabIndex        =   3
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">>"
         Height          =   375
         Left            =   5160
         TabIndex        =   17
         Top             =   2280
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   4335
         Left            =   6360
         TabIndex        =   16
         Top             =   480
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   405
         ItemData        =   "frm_update_package_details.frx":0000
         Left            =   2760
         List            =   "frm_update_package_details.frx":001C
         TabIndex        =   1
         Text            =   "(Select Sem)"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frm_update_package_details.frx":0062
         Left            =   2760
         List            =   "frm_update_package_details.frx":0072
         TabIndex        =   0
         Text            =   "(Select Field)"
         Top             =   480
         Width           =   2655
      End
      Begin VB.ComboBox Combo3 
         Height          =   405
         ItemData        =   "frm_update_package_details.frx":0092
         Left            =   2760
         List            =   "frm_update_package_details.frx":009C
         TabIndex        =   2
         Text            =   "(Select)"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Enter Subject :"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "1st Installment :"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "2nd Installment :"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "3rd Intallment :"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Price :"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Select Sem :"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Select Field :"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Select Package :"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frm_update_package_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim count1 As Integer
Dim count2 As Integer

Private Sub Combo3_Click()
    Set rs = cnn.Execute("Select * from package_details where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "' and package_type='" & Combo3.Text & "'")
    If (Not rs.EOF) Then
        Text6.Text = rs.Fields("price")
        If rs.Fields("subject1") <> "Null" Then
            List1.AddItem (rs.Fields("subject1"))
        End If
        If rs.Fields("subject2") <> "Null" Then
            List1.AddItem (rs.Fields("subject2"))
        End If
        If rs.Fields("subject3") <> "Null" Then
            List1.AddItem (rs.Fields("subject3"))
        End If
        If rs.Fields("subject4") <> "Null" Then
            List1.AddItem (rs.Fields("subject4"))
        End If
        If rs.Fields("subject5") <> "Null" Then
            List1.AddItem (rs.Fields("subject5"))
        End If
        If rs.Fields("subject6") <> "Null" Then
            List1.AddItem (rs.Fields("subject6"))
        End If
        Text9.Text = rs.Fields("first_installment")
        Text8.Text = rs.Fields("second_installment")
        Text7.Text = rs.Fields("third_installment")
    Else
        MsgBox Combo3.Text + " is not available", vbOKOnly + vbExclamation, "Information"
    End If
    If List1.ListCount <> 0 Then
        Command2.Enabled = True
    End If
End Sub

Private Sub Command1_Click()
    List1.AddItem (Text5.Text)
    Text5.Text = ""
End Sub

Private Sub Command2_Click()
    Dim a As Integer
    For i = List1.ListCount - 1 To 0 Step -1
        If List1.Selected(i) = True Then
            List1.RemoveItem (i)
        End If
    Next i
End Sub

Private Sub Command3_Click()
If Text6.Text = "" Or Text9.Text = "" Or Text8.Text = "" Or Text7.Text = "" Then
    MsgBox "Please enter all details to update", vbCritical + vbOKCancel, "warning"
Else

        s = "update package_details set price='" & Text6.Text & "',first_installment='" & Text9.Text & "',second_installment='" & Text8.Text & "',third_installment='" & Text7.Text & "' where field='" & Combo1.Text & "' and sem='" & Combo2.Text & "' and package_type='" & Combo3.Text & "'"
        cnn.Execute s
        MsgBox "Details updated successfully", vbInformation + vbOKOnly, "Successfull"
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

Private Sub Command3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command3_Click
    End If
End Sub

Private Sub Form_Load()
     Me.Top = 1000
    Me.Left = 2800
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



Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command3_Click
    End If
End Sub
