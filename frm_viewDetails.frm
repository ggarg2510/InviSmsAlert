VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form5"
   ClientHeight    =   7815
   ClientLeft      =   2760
   ClientTop       =   4305
   ClientWidth     =   12615
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_return 
      Caption         =   "&RETURN"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   6
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton btn_go 
      Caption         =   "&GO"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   5
      Top             =   5280
      Width           =   2655
   End
   Begin VB.OptionButton opt_by_detail 
      Caption         =   "By Bill Details"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7680
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.OptionButton opt_by_name 
      Caption         =   "By Customers Name"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox txt_by_detail 
      Height          =   735
      Left            =   7080
      TabIndex        =   2
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox txt_by_name 
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Enter Details ....."
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3165
      TabIndex        =   0
      Top             =   1080
      Width           =   5565
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbrecordset As Recordset
Dim db As Database
Dim dbIndex As Index


Private Sub btn_go_Click()

  Dim SearchStr As String
  Dim i As Integer
  
  If opt_by_name.Value = True Then
      SearchStr = txt_by_name.Text
  ElseIf opt_by_detail.Value = True Then
      SearchStr = txt_by_detail.Text
  End If
  

  If SearchStr <> "" Then
  
    dbrecordset.MoveFirst
    For i = 1 To dbrecordset.RecordCount
      If dbrecordset.fields(1) = SearchStr Then
        valuePasser5To3 = dbrecordset.fields(0)
        Form3.Show
        Exit Sub
      Else
        dbrecordset.MoveNext
        If dbrecordset.EOF Then
          dbrecordset.MovePrevious
        End If
      End If
    Next i
    MsgBox ("No Such Record Found ...!!")
  Else
    MsgBox ("Not Valid ...!!")
  End If

 ' Form5.Hide
  'If Form1.ViewSource = 1 Then
  '    Form1.Enabled = True
  '    Form1.Show
 ' ElseIf Form1.ViewSource = 2 Then
 ''     Form2.Enabled = True
 '     Form2.Show
 ' End If
 
End Sub

Private Sub btn_return_Click()
Form5.Hide
If Form1.ViewSource = 1 Then
    Form1.Enabled = True
    Form1.Show
ElseIf Form1.ViewSource = 2 Then
    Form2.Enabled = True
    Form2.Show
End If

End Sub

Private Sub Form_Load()
With Form5
  .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
End With
txt_by_name.Enabled = True
txt_by_detail.Enabled = False
opt_by_name.Value = True

Set db = DBEngine.Workspaces(0).OpenDatabase("db")
Set dbrecordset = db.OpenRecordset("CUSTOMERS", dbOpenTable)

End Sub

Private Sub opt_by_detail_Click()
txt_by_name.Enabled = False
txt_by_detail.Enabled = True

End Sub

Private Sub opt_by_name_Click()
txt_by_name.Enabled = True
txt_by_detail.Enabled = False

End Sub
