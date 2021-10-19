VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_pay_installment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pay Installments :"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10530
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
   ScaleHeight     =   8205
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   5040
      TabIndex        =   30
      Top             =   0
      Width           =   5295
      Begin VB.ComboBox Combo3 
         Height          =   405
         ItemData        =   "frm_pay_installment.frx":0000
         Left            =   120
         List            =   "frm_pay_installment.frx":0010
         TabIndex        =   35
         Text            =   "(Field)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frm_pay_installment.frx":0030
         Left            =   2760
         List            =   "frm_pay_installment.frx":004C
         TabIndex        =   34
         Text            =   "(Sem)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox Combo4 
         Height          =   405
         Left            =   840
         TabIndex        =   33
         Text            =   "(Batch)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   4815
      Begin VB.OptionButton Option2 
         Caption         =   "By Batch"
         Height          =   285
         Left            =   2280
         TabIndex        =   32
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Semesters"
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Installments :"
      Height          =   7575
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   10455
      Begin VB.ComboBox Combo5 
         Height          =   405
         Left            =   3480
         TabIndex        =   36
         Text            =   "(Select Name)"
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   3360
         TabIndex        =   2
         Top             =   5160
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   8040
         TabIndex        =   27
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125960193
         CurrentDate     =   43859
      End
      Begin VB.Frame Frame2 
         Caption         =   "Details :"
         Height          =   3015
         Left            =   6840
         TabIndex        =   19
         Top             =   3000
         Width           =   3495
         Begin VB.Label Label16 
            Height          =   375
            Left            =   2040
            TabIndex        =   25
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "3rd Installment :"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label14 
            Height          =   375
            Left            =   2040
            TabIndex        =   23
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "2nd Installment :"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label12 
            Height          =   375
            Left            =   2040
            TabIndex        =   21
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "1st Installment :"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   3360
         TabIndex        =   3
         Top             =   5880
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Pay"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   6720
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frm_pay_installment.frx":0092
         Left            =   3360
         List            =   "frm_pay_installment.frx":0094
         TabIndex        =   1
         Text            =   "(Select)"
         Top             =   4440
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   3360
         TabIndex        =   14
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   3360
         TabIndex        =   12
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   3360
         TabIndex        =   10
         Top             =   2130
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   7080
         TabIndex        =   0
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Enter Receipt No :"
         Height          =   375
         Left            =   600
         TabIndex        =   28
         Top             =   5160
         Width           =   2175
      End
      Begin VB.Label Label17 
         Caption         =   "Date :"
         Height          =   375
         Left            =   6960
         TabIndex        =   26
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   495
         Left            =   9000
         TabIndex        =   18
         Top             =   7080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   615
         Left            =   7560
         TabIndex        =   17
         Top             =   7080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Enter Amount :"
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Select Installment :"
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Balance Fees :"
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Package Type :"
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Student Name :"
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Student RollNo :"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Rollno or Name :"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frm_pay_installment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
Dim f As Integer
Dim count1 As Integer
Dim count2 As Integer

Private Sub Combo2_Click()
    Combo4.Clear
    If Combo3.Text = "(Field)" Then
        MsgBox "Please select the Field", vbQuestion + vbOKOnly, "Warning"
    ElseIf Combo2.Text = "(Sem)" Then
        MsgBox "Please select the Semester", vbQuestion + vbOKOnly, "Warning"
    Else
        Set rs = cnn.Execute("select * from admission where field='" & Combo3.Text & "' and sem='" & Combo2.Text & "'")
        Do While Not rs.EOF
            Combo5.AddItem (rs.Fields("sname"))
            rs.MoveNext
        Loop
    End If
End Sub

Private Sub Combo3_Click()
    If Combo3.Text = "OTHER" Then
        If count2 = 0 And Combo3.Text = "OTHER" Then
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

    Combo4.Clear
    If Combo2.Text <> "(Sem)" Then
        Set rs = cnn.Execute("select * from admission where field='" & Combo3.Text & "' and sem='" & Combo2.Text & "'")
        Do While Not rs.EOF
            Combo5.AddItem (rs.Fields("sname"))
            rs.MoveNext
        Loop
    End If
End Sub


Private Sub Combo4_Click()
    Set rs = cnn.Execute("Select * from stud_batch_details where batch='" & Combo4.Text & "'")
    Do While Not rs.EOF
        Combo5.AddItem (rs.Fields("sname"))
        rs.MoveNext
    Loop
End Sub



Private Sub Combo5_Click()
     If Combo5.Text = "(Select Name)" Then
        MsgBox "Please select the name of student", vbQuestion + vbOKOnly, "Warning"
    Else
        Set rs = cnn.Execute("select * from admission where sname='" & Combo5.Text & "'")
        If (Not rs.EOF) Then
            Label3.Caption = rs.Fields("rollno")
            Text2.Text = rs.Fields("sname")
            Text3.Text = rs.Fields("subject")
            Label9.Caption = rs.Fields("field")
            Label10.Caption = rs.Fields("sem")
            If Text3.Text = "Select Subjects" Then
                Set rs = cnn.Execute("select * from trans where sname='" & Combo5.Text & "'")
                If (Not rs.EOF) Then
                    If rs.Fields("twoinstallment") <> "Null" Then
                        d = Val(rs.Fields("firstinstallment"))
                        e = Val(rs.Fields("secondinstallment"))
                        Label15.Visible = False
                        Label16.Visible = False
                        a = Val(rs.Fields("oneinstallment"))
                        b = Val(rs.Fields("twoinstallment"))
                        Text4.Text = rs.Fields("balance")
                    Else
                        d = Val(rs.Fields("firstinstallment"))
                        a = Val(rs.Fields("oneinstallment"))
                        Text4.Text = rs.Fields("balance")
                        Label13.Visible = False
                        Label14.Visible = False
                        Label15.Visible = False
                        Label16.Visible = False
                    End If
                End If
                
            Else
                Set rs = cnn.Execute("select * from trans where sname='" & Combo5.Text & "'")
                If (Not rs.EOF) Then
                    Text4.Text = rs.Fields("balance")
                    d = Val(rs.Fields("firstinstallment"))
                    e = Val(rs.Fields("secondinstallment"))
                    f = Val(rs.Fields("thirdinstallment"))
                End If
    
                Set rs = cnn.Execute("select * from package_details where field='" & Label9.Caption & "' and sem='" & Label10.Caption & "' and package_type='" & Text3.Text & "'")
                If (Not rs.EOF) Then
                    a = Val(rs.Fields("first_installment"))
                    b = Val(rs.Fields("second_installment"))
                    c = Val(rs.Fields("third_installment"))
                    If c = "0" Then
                        Label15.Visible = False
                        Label16.Visible = False
                    End If
                End If
             End If
                If a > d Then
                    Label12.Caption = a - d
                ElseIf a = d Then
                    Label12.Caption = "Paid"
                End If
                If b > e Then
                    Label14.Caption = b - e
                ElseIf b = e Then
                    Label14.Caption = "Paid"
                End If
                If c = 0 Then
                    Label16.Caption = "N.A."
                ElseIf c > f Then
                    Label16.Caption = c - f
                ElseIf c = f Then
                    Label16.Caption = "Paid"
                    MsgBox "Installments are already paid,  Thank You ", vbInformation + vbOKOnly, "Information"
                        Text2.Text = ""
                        Text3.Text = ""
                        Text4.Text = ""
                        Text5.Text = ""
                        Label3.Caption = ""
                        Combo1.Clear
                        Label12.Caption = ""
                        Label14.Caption = ""
                        Label16.Caption = ""
                        Text6.Text = ""
                        Command2.Enabled = False
                        
                    Exit Sub
                End If
           
            
            If Label12.Caption <> "Paid" Then
                Combo1.AddItem ("1st Installment")
            ElseIf Label14.Caption <> "Paid" Then
                Combo1.AddItem ("2nd Installment")
            ElseIf Label16.Caption <> "Paid" Then
                Combo1.AddItem ("3rd Installment")
            End If
            Command2.Enabled = True
        Else
            MsgBox "Student is not found", vbExclamation + vbOKOnly, "Informatioin"
        End If
    End If
    If Text4.Text = "0" Then
        MsgBox "Installments are already paid", vbInformation + vbOKOnly, "Information"
        Command2.Enabled = False
                        Text2.Text = ""
                        Text3.Text = ""
                        Text4.Text = ""
                        Text5.Text = ""
                        Label3.Caption = ""
                        Combo1.Clear
                        Label12.Caption = ""
                        Label14.Caption = ""
                        Label16.Caption = ""
                        Text6.Text = ""
                        Command2.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim a As Integer
    a = Val(Text5.Text)
    Dim res As VbMsgBoxResult
    Dim k As String
    Dim r As String
    Dim bal As String
    If Text4.Text <> 0 Then
        If Combo1.Text = "1st Installment" Then
            If Label12.Caption <> "Paid" Then
                If Val(Label12.Caption) < Val(Text5.Text) Then
                    MsgBox "first complete current installment then jump next one", vbCritical + vbOKOnly, "Warning"
                Else
                    If Val(Label12.Caption) > Val(Text5.Text) Then
                        res = MsgBox("Amount is less than the remaining installment,are you sure to proceed ", vbYesNo + vbCritical, "Warning")
                        If res = vbYes Then
                            k = Val(Text5.Text) + d
                            r = Val(Label12.Caption) - Val(Text5.Text)
                            bal = Val(Text4.Text) - Val(Text5.Text)
                            s = "update trans set firstinstallment='" & k & "',firstdate='" & DTPicker1.Value & "',balance='" & bal & "',recieptno='" & Text6.Text & "' where rollno='" & Label3.Caption & "' and sname='" & Text2.Text & "'"
                            cnn.Execute s
                            m = "insert into bill_view values ('" + Text2.Text + "','" + Text6.Text + "','" & a & "','" & DTPicker1.Value & "')"
                            cnn.Execute m
                            MsgBox "1st Installment paid successfully remaining = " + r, vbInformation + vbOKOnly, "Information"
                            Text2.Text = ""
                            Text3.Text = ""
                            Text4.Text = ""
                            Text5.Text = ""
                            Label3.Caption = ""
                            Combo1.Clear
                            Label12.Caption = ""
                            Label14.Caption = ""
                            Label16.Caption = ""
                            Text6.Text = ""
                            Command2.Enabled = False
                            Unload Me
                        End If
                    Else
                            k = Val(Text5.Text) + d
                            bal = Val(Text4.Text) - Val(Text5.Text)
                            s = "update trans set firstinstallment='" & k & "',firstdate='" & DTPicker1.Value & "',balance='" & bal & "',recieptno='" & Text6.Text & "' where rollno='" & Label3.Caption & "' and sname='" & Text2.Text & "'"
                            cnn.Execute s
                            m = "insert into bill_view values ('" + Text2.Text + "','" + Text6.Text + "','" & a & "','" & DTPicker1.Value & "')"
                            cnn.Execute m
                            MsgBox "1st Installment paid successfully", vbInformation + vbOKOnly, "Information"
                        
                            Text2.Text = ""
                            Text3.Text = ""
                            Text4.Text = ""
                            Text5.Text = ""
                            Label3.Caption = ""
                            Combo1.Clear
                            Label12.Caption = ""
                            Label14.Caption = ""
                            Label16.Caption = ""
                            Text6.Text = ""
                            Command2.Enabled = False
                            Unload Me
                    End If
                End If
            End If
        ElseIf Combo1.Text = "2nd Installment" Then
            If Label14.Caption <> "Paid" Then
                If Val(Label14.Caption) < Val(Text5.Text) Then
                    MsgBox "first complete current installment then jump next one", vbCritical + vbOKOnly, "Warning"
                Else
                    If Val(Label14.Caption) > Val(Text5.Text) Then
                        res = MsgBox("Amount is less than the remaining installment,are you sure to proceed ", vbYesNo + vbCritical, "Warning")
                        If res = vbYes Then
                            k = Val(Text5.Text) + e
                            r = Val(Label14.Caption) - Val(Text5.Text)
                            bal = Val(Text4.Text) - Val(Text5.Text)
                            s = "update trans set secondinstallment='" & k & "',seconddate='" & DTPicker1.Value & "',balance='" & bal & "',recieptno='" & Text6.Text & "' where rollno='" & Label3.Caption & "' and sname='" & Text2.Text & "'"
                            cnn.Execute s
                            m = "insert into bill_view values ('" + Text2.Text + "','" + Text6.Text + "','" & a & "','" & DTPicker1.Value & "')"
                            cnn.Execute m
                            MsgBox "2nd Installment paid successfully remaining = " + r, vbInformation + vbOKOnly, "Information"
                       
                            Text2.Text = ""
                            Text3.Text = ""
                            Text4.Text = ""
                            Text5.Text = ""
                            Label3.Caption = ""
                            Combo1.Clear
                            Label12.Caption = ""
                            Label14.Caption = ""
                            Label16.Caption = ""
                            Text6.Text = ""
                            Command2.Enabled = False
                            Unload Me
                        End If
                    Else
                            k = Val(Text5.Text) + e
                            bal = Val(Text4.Text) - Val(Text5.Text)
                            s = "update trans set secondinstallment='" & k & "',seconddate='" & DTPicker1.Value & "',balance='" & bal & "',recieptno='" & Text6.Text & "' where rollno='" & Label3.Caption & "' and sname='" & Text2.Text & "'"
                            cnn.Execute s
                            m = "insert into bill_view values ('" + Text2.Text + "','" + Text6.Text + "','" & a & "','" & DTPicker1.Value & "')"
                            cnn.Execute m
                            MsgBox "2nd Installment paid successfully", vbInformation + vbOKOnly, "Information"
                       
                            Text2.Text = ""
                            Text3.Text = ""
                            Text4.Text = ""
                            Text5.Text = ""
                            Label3.Caption = ""
                            Combo1.Clear
                            Label12.Caption = ""
                            Label14.Caption = ""
                            Label16.Caption = ""
                            Text6.Text = ""
                            Command2.Enabled = False
                            Unload Me
                    End If
                End If
            End If
        ElseIf Combo1.Text = "3rd Installment" Then
            If Label16.Caption <> "Paid" Then
                If Val(Label16.Caption) < Val(Text5.Text) Then
                    MsgBox "Money you entered is greater than the remaining installment", vbCritical + vbOKOnly, "Warning"
                Else
                    If Val(Label16.Caption) > Val(Text5.Text) Then
                        res = MsgBox("Amount is less than the remaining installment,are you sure to proceed ", vbYesNo + vbCritical, "Warning")
                        If res = vbYes Then
                            k = Val(Text5.Text) + f
                            r = Val(Label16.Caption) - Val(Text5.Text)
                            bal = Val(Text4.Text) - Val(Text5.Text)
                            s = "update trans set thirdinstallment='" & k & "',thirddate='" & DTPicker1.Value & "',balance='" & bal & "',recieptno='" & Text6.Text & "' where rollno='" & Label3.Caption & "' and sname='" & Text2.Text & "'"
                            cnn.Execute s
                            m = "insert into bill_view values ('" + Text2.Text + "','" + Text6.Text + "','" & a & "','" & DTPicker1.Value & "')"
                            cnn.Execute m
                            MsgBox "3rd Installment paid successfully remaining = " + r, vbInformation + vbOKOnly, "Information"
                       
                            Text2.Text = ""
                            Text3.Text = ""
                            Text4.Text = ""
                            Text5.Text = ""
                            Label3.Caption = ""
                            Combo1.Clear
                            Label12.Caption = ""
                            Label14.Caption = ""
                            Label16.Caption = ""
                            Text6.Text = ""
                            Command2.Enabled = False
                            Unload Me
                        End If
                    Else
                            k = Val(Text5.Text) + f
                            bal = Val(Text4.Text) - Val(Text5.Text)
                            s = "update trans set thirdinstallment='" & k & "',thirddate='" & DTPicker1.Value & "',balance='" & bal & "',recieptno='" & Text6.Text & "' where rollno='" & Label3.Caption & "' and sname='" & Text2.Text & "'"
                            cnn.Execute s
                            m = "insert into bill_view values ('" + Text2.Text + "','" + Text6.Text + "','" & a & "','" & DTPicker1.Value & "')"
                            cnn.Execute m
                            MsgBox "3rd Installment paid successfully", vbInformation + vbOKOnly, "Information"
                       
                            Text2.Text = ""
                            Text3.Text = ""
                            Text4.Text = ""
                            Text5.Text = ""
                            Label3.Caption = ""
                            Combo1.Clear
                            Label12.Caption = ""
                            Label14.Caption = ""
                            Label16.Caption = ""
                            Text6.Text = ""
                            Command2.Enabled = False
                            Unload Me
                            Exit Sub
                    End If
                End If
            ElseIf Label16.Caption = "Paid" Then
                MsgBox "Installments are already paid,  Thank You ", vbInformation + vbOKOnly, "Information"
                Unload Me
                Exit Sub
            ElseIf Label16.Caption = "N.A." Then
                MsgBox "3rd installment is not availabel", vbInformation + vbOKOnly, "Infromation"
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    Me.Top = 600
    Me.Left = 2500
    count1 = 0
    count2 = 1
    connect
     Set rs = cnn.Execute("select * from batch_details")
    If (Not rs.EOF) Then
        Do While Not rs.EOF
            Combo4.AddItem (rs.Fields("batch"))
            rs.MoveNext
        Loop
    End If
End Sub

Private Sub Option1_Click()
    Combo3.Visible = True
    Combo2.Visible = True
    Combo4.Visible = False
End Sub

Private Sub Option2_Click()
     Combo3.Visible = False
    Combo2.Visible = False
    Combo4.Visible = True
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
Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command2_Click
    End If
End Sub
