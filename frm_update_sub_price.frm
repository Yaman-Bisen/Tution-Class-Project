VERSION 5.00
Begin VB.Form frm_update_sub_price 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Subject Details :"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6855
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
   ScaleHeight     =   4680
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Subject Details :"
      Height          =   4695
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton Command1 
         Caption         =   "&Update"
         Height          =   495
         Left            =   1680
         TabIndex        =   4
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   2760
         TabIndex        =   3
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         Height          =   405
         Left            =   2760
         TabIndex        =   2
         Text            =   "(Select Subject)"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "frm_update_sub_price.frx":0000
         Left            =   2760
         List            =   "frm_update_sub_price.frx":0010
         TabIndex        =   0
         Text            =   "(Select Field)"
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "frm_update_sub_price.frx":0030
         Left            =   2760
         List            =   "frm_update_sub_price.frx":004C
         TabIndex        =   1
         Text            =   "(Select Sem)"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "New Price :"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label5 
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Old Price :"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Select Subject :"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Select Field :"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Select Sem :"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_update_sub_price"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim count1 As Integer
Dim count2 As Integer

Private Sub Combo2_Click()
    Combo3.Clear
    If Combo2.Text = "OTHER" Then
        Set rs = cnn.Execute("select distinct subject from subject_details order by subject")
        Do While Not rs.EOF
            Combo3.AddItem (rs.Fields("subject"))
            rs.MoveNext
        Loop
    Else
        Set rs = cnn.Execute("Select * from subject_details where years='" + Combo1.Text + "' and sem='" + Combo2.Text + "'")
        If (Not rs.EOF) Then
            Do While Not rs.EOF
                Combo3.AddItem (rs.Fields("subject"))
                rs.MoveNext
            Loop
        Else
            MsgBox "Subjects are not availabel for the selected information", vbInformation + vbOKOnly, "Information"
        End If
    End If
End Sub

Private Sub Combo3_Click()
    Set rs = cnn.Execute("select * from subject_details where years='" + Combo1.Text + "' and sem='" + Combo2.Text + "' and subject='" + Combo3.Text + "'")
    If (Not rs.EOF) Then
        Label5.Caption = rs.Fields("cost")
    End If
End Sub

Private Sub Command1_Click()
    If Combo1.Text = "(Select Field)" Then
        MsgBox "Please select the field", vbExclamation + vbOKOnly, "Warning"
    ElseIf Combo2.Text = "(Select Sem)" Then
        MsgBox "Please select semester", vbExclamation + vbOKOnly, "Warning"
    ElseIf Combo3.Text = "(Select Subject)" Then
        MsgBox "Please select the Subject", vbExclamation + vbOKOnly, "Warning"
    ElseIf Text1.Text = "" Then
        MsgBox "Please enter the price to update", vbExclamation + vbOKOnly, "Warning"
    Else
        s = "update subject_details set cost='" & Text1.Text & "' where subject='" & Combo3.Text & "' and years='" & Combo1.Text & "' and sem='" & Combo2.Text & "'"
        cnn.Execute s
        MsgBox "Subject price updated successfully", vbInformation + vbOKOnly, "Information"
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
    connect
    count1 = 0
    count2 = 1
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
