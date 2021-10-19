VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   0
      List            =   "Form1.frx":000A
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   0
      Width           =   3375
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1920
      Width           =   3375
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   3600
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As New ADODB.Recordset
Private Sub Combo1_Click()
    If Combo1.Text = "By Student Name" Then
        Set rs = cnn.Execute("Select * from admission")
        Do While Not rs.EOF
            Combo2.AddItem (rs.Fields(3))
            rs.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Load()
    connect
End Sub


