VERSION 5.00
Begin VB.Form frm_transaction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pay 1st Intallment :"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11055
   ControlBox      =   0   'False
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
   ScaleHeight     =   6480
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Installment :"
      Height          =   5295
      Left            =   6000
      TabIndex        =   15
      Top             =   1080
      Width           =   4935
      Begin VB.TextBox Text7 
         Height          =   405
         Left            =   2880
         TabIndex        =   0
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   2880
         TabIndex        =   1
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Pay"
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "Enter Receipt No :"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "Pay 1st installment :"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   375
         Left            =   2880
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   375
         Left            =   2880
         TabIndex        =   20
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   375
         Left            =   2880
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Third Installment :"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Second Installment :"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "First Installment :"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      Height          =   375
      Left            =   480
      TabIndex        =   23
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Package Price :"
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Student Name :"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Package :"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Semester :"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Field :"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Roll No :"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Roll No :"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frm_transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
    Dim res As VbMsgBoxResult
    If Text6.Text = "" Then
        MsgBox "please enter the amount to pay ", vbCritical + vbOKOnly, "Warning"
    Else
        If Text4.Text = "Select Subjects" Then
            If count5 = "1" Then
                If Val(Label11.Caption) > Val(Text6.Text) Then
                    res = MsgBox("Amount is less than first installment,Are you sure to proceed", vbQuestion + vbYesNo, "Information")
                    If res = vbYes Then
                        Label16.Caption = Val(Text5.Text) - Val(Text6.Text)
                        s = "insert into trans (rollno,sname,firstdate,balance,firstinstallment,secondinstallment,seconddate,thirdinstallment,thirddate,oneinstallment,twoinstallment,recieptno) values ('" & rollno1 & "','" & Text1.Text & "','" & date1 & "','" & Label16.Caption & "','" & Text6.Text & "','" & "0" & "','" & "0" & "','" & "0" & "','" & "0" & "','" & price1 & "','" & "Null" & "','" & Text7.Text & "')"
                        cnn.Execute s
                        MsgBox "Amount paid successfully", vbOKOnly + vbInformation, "Successfull"
                        m = "insert into bill_view values ('" + Text1.Text + "','" + Text7.Text + "','" + Text6.Text + "','" & date1 & "')"
                        cnn.Execute m
                        Unload Me
                    End If
                Else
                    Label16.Caption = Val(Text5.Text) - Val(Text6.Text)
                    s = "insert into trans (rollno,sname,firstdate,balance,firstinstallment,secondinstallment,seconddate,thirdinstallment,thirddate,oneinstallment,twoinstallment,recieptno) values ('" & rollno1 & "','" & Text1.Text & "','" & date1 & "','" & Label16.Caption & "','" & Text6.Text & "','" & "0" & "','" & "0" & "','" & "0" & "','" & "0" & "','" & price1 & "','" & "Null" & "','" & Text7.Text & "')"
                    cnn.Execute s
                    MsgBox "Amount paid successfully", vbOKOnly + vbInformation, "Successfull"
                    m = "insert into bill_view values ('" + Text1.Text + "','" + Text7.Text + "','" + Text6.Text + "','" & date1 & "')"
                    cnn.Execute m
                    Unload Me
                End If
            Else
                If Val(Label11.Caption) > Val(Text6.Text) Then
                    res = MsgBox("Amount is less than first installment,Are you sure to proceed", vbQuestion + vbYesNo, "Information")
                    If res = vbYes Then
                        Label16.Caption = Val(Text5.Text) - Val(Text6.Text)
                        s = "insert into trans (rollno,sname,firstdate,balance,firstinstallment,secondinstallment,seconddate,thirdinstallment,thirddate,oneinstallment,twoinstallment,recieptno) values ('" & rollno1 & "','" & Text1.Text & "','" & date1 & "','" & Label16.Caption & "','" & Text6.Text & "','" & "0" & "','" & "0" & "','" & "0" & "','" & "0" & "','" & price1 & "','" & price2 & "','" & Text7.Text & "')"
                        cnn.Execute s
                        MsgBox "Amount paid successfully", vbOKOnly + vbInformation, "Successfull"
                        m = "insert into bill_view values ('" + Text1.Text + "','" + Text7.Text + "','" + Text6.Text + "','" & date1 & "')"
                        cnn.Execute m
                        Unload Me
                    End If
                Else
                    Label16.Caption = Val(Text5.Text) - Val(Text6.Text)
                    s = "insert into trans (rollno,sname,firstdate,balance,firstinstallment,secondinstallment,seconddate,thirdinstallment,thirddate,oneinstallment,twoinstallment,recieptno) values ('" & rollno1 & "','" & Text1.Text & "','" & date1 & "','" & Label16.Caption & "','" & Text6.Text & "','" & "0" & "','" & "0" & "','" & "0" & "','" & "0" & "','" & price1 & "','" & price2 & "','" & Text7.Text & "')"
                    cnn.Execute s
                    MsgBox "Amount paid successfully", vbOKOnly + vbInformation, "Successfull"
                    m = "insert into bill_view values ('" + Text1.Text + "','" + Text7.Text + "','" + Text6.Text + "','" & date1 & "')"
                    cnn.Execute m
                    Unload Me
                End If
            End If
        Else
            If Val(Label11.Caption) > Val(Text6.Text) Then
                res = MsgBox("Amount is less than first installment,Are you sure to proceed", vbQuestion + vbYesNo, "Information")
                If res = vbYes Then
                    Label16.Caption = Val(Text5.Text) - Val(Text6.Text)
                    s = "insert into trans (rollno,sname,firstdate,balance,firstinstallment,secondinstallment,seconddate,thirdinstallment,thirddate,recieptno) values ('" & rollno1 & "','" & Text1.Text & "','" & date1 & "','" & Label16.Caption & "','" & Text6.Text & "','" & "0" & "','" & "0" & "','" & "0" & "','" & "0" & "','" & Text7.Text & "')"
                    cnn.Execute s
                    m = "insert into bill_view values ('" + Text1.Text + "','" + Text7.Text + "','" + Text6.Text + "','" & date1 & "')"
                    cnn.Execute m
                    MsgBox "Amount paid successfully", vbOKOnly + vbInformation, "Successfull"
                    Unload Me
                    Close
                End If
            Else
                Label16.Caption = Val(Text5.Text) - Val(Text6.Text)
                s = "insert into trans (rollno,sname,firstdate,balance,firstinstallment,secondinstallment,seconddate,thirdinstallment,thirddate,recieptno) values ('" & rollno1 & "','" & Text1.Text & "','" & date1 & "','" & Label16.Caption & "','" & Text6.Text & "','" & "0" & "','" & "0" & "','" & "0" & "','" & "0" & "','" & Text7.Text & "')"
                cnn.Execute s
                MsgBox "Amount paid successfully", vbOKOnly + vbInformation, "Successfull"
                m = "insert into bill_view values ('" + Text1.Text + "','" + Text7.Text + "','" + Text6.Text + "','" & date1 & "')"
                cnn.Execute m
                Unload Me
            End If
        End If
    End If
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Form_Load()
     Me.Top = 1500
    Me.Left = 2500
    connect
    Label2.Caption = rollno1
    Text1.Text = sname1
    Text2.Text = field
    Text3.Text = sem
    Text4.Text = package_type
    Text5.Text = cost
    
    If package_type = "Select Subjects" Then
        If count5 = "1" Then
            
            Label8.Visible = True
            Label11.Visible = True
            Label9.Visible = False
            Label10.Visible = False
            Label12.Visible = False
            Label13.Visible = False
            Label11.Caption = price1
        ElseIf count5 = "2" Then
            Label8.Visible = True
            Label11.Visible = True
            Label9.Visible = True
            Label10.Visible = False
            Label12.Visible = True
            Label13.Visible = False
            Label11.Caption = price1
            Label12.Caption = price2
        End If
        
    Else
            Label8.Visible = True
            Label11.Visible = True
            Label9.Visible = True
            Label10.Visible = True
            Label12.Visible = True
            Label13.Visible = True
        Set rs = cnn.Execute("select * from package_details where field='" & field & "' and sem='" & sem & "' and package_type='" & package_type & "'")
        If (Not rs.EOF) Then
            Label11.Caption = rs.Fields("first_installment")
            Label12.Caption = rs.Fields("second_installment")
            Label13.Caption = rs.Fields("third_installment")
            If Label13.Caption = "0" Then
                Label10.Visible = False
                Label13.Visible = False
            End If
        End If

    End If
End Sub


Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub
