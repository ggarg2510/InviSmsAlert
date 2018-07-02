VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   Caption         =   "Form3"
   ClientHeight    =   9705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14280
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11988.53
   ScaleMode       =   0  'User
   ScaleWidth      =   14430
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
      Height          =   975
      Left            =   8760
      TabIndex        =   26
      Top             =   8040
      Width           =   3255
   End
   Begin VB.CommandButton btn_delete 
      Caption         =   "&DELETE RECORD"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      TabIndex        =   25
      Top             =   8040
      Width           =   3255
   End
   Begin VB.CommandButton btn_prev 
      Caption         =   "<"
      Height          =   495
      Left            =   7080
      TabIndex        =   24
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton btn_next 
      Caption         =   ">"
      Height          =   495
      Left            =   8520
      TabIndex        =   23
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton btn_last 
      Caption         =   ">>"
      Height          =   495
      Left            =   9960
      TabIndex        =   22
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton btn_first 
      Caption         =   "<<"
      Height          =   495
      Left            =   5640
      TabIndex        =   21
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox add1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1920
      TabIndex        =   8
      Text            =   "1234567890123456789012345678901234567890"
      Top             =   4080
      Width           =   6735
   End
   Begin VB.TextBox add2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1920
      TabIndex        =   7
      Text            =   "1234567890123456789012345678901234567890"
      Top             =   5040
      Width           =   6735
   End
   Begin VB.TextBox mob1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1920
      MaxLength       =   12
      TabIndex        =   6
      Text            =   "123456789012345"
      Top             =   6000
      Width           =   2655
   End
   Begin VB.TextBox mob2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1920
      MaxLength       =   12
      TabIndex        =   5
      Text            =   "123456789012345"
      Top             =   6960
      Width           =   2655
   End
   Begin VB.TextBox billDetail 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   1920
      TabIndex        =   3
      Text            =   "1234567890123456789012345678901234567890"
      Top             =   3120
      Width           =   6735
   End
   Begin VB.TextBox CustName 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1920
      MaxLength       =   32
      TabIndex        =   2
      Text            =   "12345678901234567890123456789012"
      Top             =   2160
      Width           =   5535
   End
   Begin VB.TextBox recordNo 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "1"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox NxtAlert 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10680
      MaxLength       =   12
      TabIndex        =   0
      Text            =   "12"
      Top             =   6000
      Width           =   735
   End
   Begin MSComCtl2.MonthView purchaseDate 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-MM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      Height          =   2370
      Left            =   10200
      TabIndex        =   4
      Top             =   2160
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483642
      Appearance      =   1
      Enabled         =   0   'False
      OLEDropMode     =   1
      StartOfWeek     =   43515905
      TitleBackColor  =   -2147483627
      TitleForeColor  =   -2147483633
      TrailingForeColor=   -2147483636
      CurrentDate     =   42099
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "READ RECORD"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4935
      TabIndex        =   19
      Top             =   0
      Width           =   3825
   End
   Begin VB.Label Label2 
      Caption         =   "Record #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Address 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Address 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Mobile 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Mobile 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Date of Purchase"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   13
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label9 
      Caption         =   "Bill Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lbl_date 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "12/12/12"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11055
      TabIndex        =   11
      Top             =   1680
      Width           =   945
   End
   Begin VB.Label Label10 
      Caption         =   "Alert After Every"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   10
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label Label11 
      Caption         =   "Months"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      TabIndex        =   9
      Top             =   6000
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbrecordset As Recordset
Dim db As Database
Dim dbIndex As Index

Sub updateReadRec()
  Dim mmByName(12) As String
  mmByName(1) = "JAN"
  mmByName(2) = "FEB"
  mmByName(3) = "MAR"
  mmByName(4) = "APR"
  mmByName(5) = "MAY"
  mmByName(6) = "JUN"
  mmByName(7) = "JUL"
  mmByName(8) = "AUG"
  mmByName(9) = "SEP"
  mmByName(10) = "OCT"
  mmByName(11) = "NOV"
  mmByName(12) = "DEC"
    
  recordNo = dbrecordset.fields(0)
  CustName.Text = dbrecordset.fields(1)
  billDetail.Text = dbrecordset.fields(2)
  add1.Text = dbrecordset.fields(3)
  add2.Text = dbrecordset.fields(4)
  mob1.Text = dbrecordset.fields(5)
  mob2.Text = dbrecordset.fields(6)
  purchaseDate.Value = dbrecordset.fields(7)
  NxtAlert.Text = dbrecordset.fields(11)
  lbl_date.Caption = purchaseDate.Day & "-" & mmByName(Val(purchaseDate.Month)) & "-" & purchaseDate.Year

End Sub

Private Sub btn_delete_Click()
  Dim i As Byte
  i = 0
  dbrecordset.Delete
  dbrecordset.MoveNext
  If dbrecordset.EOF Then
    dbrecordset.MovePrevious
  End If
  
  While dbrecordset.EOF = False
    dbrecordset.Edit
    dbrecordset.fields(0) = dbrecordset.fields(0) - 1
    dbrecordset.Update
    dbrecordset.MoveNext
    i = i + 1
  Wend
  
  While i <> 0
    dbrecordset.MovePrevious
    i = i - 1
  Wend
  Call updateReadRec
  
End Sub

Private Sub btn_first_Click()
  dbrecordset.MoveFirst
  Call updateReadRec
End Sub

Private Sub btn_last_Click()
  dbrecordset.MoveLast
  Call updateReadRec
End Sub

Private Sub btn_next_Click()
  dbrecordset.MoveNext
  If dbrecordset.EOF Then
    dbrecordset.MovePrevious
  End If
  Call updateReadRec
End Sub

Private Sub btn_prev_Click()
  dbrecordset.MovePrevious
  If dbrecordset.BOF Then
      dbrecordset.MoveNext
  End If
  Call updateReadRec
End Sub

Private Sub btn_return_Click()
Form3.Hide
Form5.Show

End Sub

Private Sub Form_Load()
  Dim i As Integer
  
  With Form3
    .Left = Form2.Left
    .Width = Form2.Width
    .ScaleHeight = Form2.ScaleHeight
    .ScaleLeft = Form2.ScaleLeft
    .ScaleMode = Form2.ScaleMode
    .ScaleTop = Form2.ScaleTop
    .ScaleWidth = Form2.ScaleWidth
    
    .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
  End With
  
  recordNo = 0
  CustName.Text = ""
  billDetail.Text = ""
  add1.Text = ""
  add2.Text = ""
  mob1.Text = ""
  mob2.Text = ""
  purchaseDate.Value = 0
  NxtAlert.Text = ""
  
  Set db = DBEngine.Workspaces(0).OpenDatabase("db")
  Set dbrecordset = db.OpenRecordset("CUSTOMERS", dbOpenTable)


  dbrecordset.MoveFirst
  For i = 1 To dbrecordset.RecordCount
    If dbrecordset.fields(0) = valuePasser5To3 Then
      Call updateReadRec
      Exit For
    Else
      dbrecordset.MoveNext
      If dbrecordset.EOF Then
        dbrecordset.MovePrevious
      End If
    End If
  Next i
      
End Sub
