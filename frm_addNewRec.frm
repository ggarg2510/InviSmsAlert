VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   9135
   ClientLeft      =   9180
   ClientTop       =   780
   ClientWidth     =   14430
   FillStyle       =   4  'Upward Diagonal
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   14430
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
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   8
      Text            =   "12"
      Top             =   6120
      Width           =   735
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
      Left            =   2280
      TabIndex        =   23
      Text            =   "1"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox CustName 
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
      Left            =   2280
      MaxLength       =   32
      TabIndex        =   1
      Text            =   "12345678901234567890123456789012"
      Top             =   2280
      Width           =   5535
   End
   Begin VB.CommandButton btn_Back 
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
      Left            =   8520
      TabIndex        =   11
      Top             =   7800
      Width           =   3255
   End
   Begin VB.TextBox billDetail 
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
      Left            =   2280
      TabIndex        =   2
      Text            =   "1234567890123456789012345678901234567890"
      Top             =   3240
      Width           =   6735
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
      Left            =   10560
      TabIndex        =   7
      Top             =   2280
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483642
      Appearance      =   1
      OLEDropMode     =   1
      StartOfWeek     =   103546881
      TitleBackColor  =   -2147483627
      TitleForeColor  =   -2147483633
      TrailingForeColor=   -2147483636
      CurrentDate     =   42099
   End
   Begin VB.CommandButton btn_ClrAll 
      Caption         =   "&CLEAR FIELDS"
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
      Left            =   4440
      TabIndex        =   10
      Top             =   7800
      Width           =   3255
   End
   Begin VB.CommandButton btn_Updt 
      Caption         =   "&UPDATE RECORD"
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
      Left            =   360
      TabIndex        =   9
      Top             =   7800
      Width           =   3255
   End
   Begin VB.TextBox mob2 
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
      Left            =   2280
      MaxLength       =   12
      TabIndex        =   6
      Text            =   "123456789012345"
      Top             =   7080
      Width           =   2655
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
      Left            =   2280
      MaxLength       =   12
      TabIndex        =   5
      Text            =   "123456789012345"
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox add2 
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
      Left            =   2280
      TabIndex        =   4
      Text            =   "1234567890123456789012345678901234567890"
      Top             =   5160
      Width           =   6735
   End
   Begin VB.TextBox add1 
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
      Left            =   2280
      TabIndex        =   3
      Text            =   "1234567890123456789012345678901234567890"
      Top             =   4200
      Width           =   6735
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
      Left            =   12000
      TabIndex        =   22
      Top             =   6120
      Width           =   1095
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
      Left            =   10800
      TabIndex        =   21
      Top             =   5520
      Width           =   2415
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
      Left            =   11415
      TabIndex        =   20
      Top             =   1800
      Width           =   945
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
      Left            =   360
      TabIndex        =   19
      Top             =   3240
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
      Left            =   10440
      TabIndex        =   18
      Top             =   1200
      Width           =   2895
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
      Left            =   360
      TabIndex        =   17
      Top             =   7080
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
      Left            =   360
      TabIndex        =   16
      Top             =   6120
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
      Left            =   360
      TabIndex        =   15
      Top             =   5160
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
      Left            =   360
      TabIndex        =   14
      Top             =   4200
      Width           =   1575
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
      Left            =   360
      TabIndex        =   13
      Top             =   2280
      Width           =   1575
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
      Left            =   360
      TabIndex        =   12
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "ADD NEW RECORD"
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
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuViewRecord 
         Caption         =   "View Record"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuDelRecord 
         Caption         =   "Delete Record"
         Shortcut        =   ^D
      End
      Begin VB.Menu h 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuRegister 
      Caption         =   "&Register"
   End
   Begin VB.Menu mnuLicense 
      Caption         =   "&License"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim dbIndex As Index

Dim dbrecordset As Recordset
Dim fields(12) As Field
Dim mmByName(12) As String
  

Private Sub clrContents()
  CustName.Text = ""
  billDetail.Text = ""
  add1.Text = ""
  add2.Text = ""
  mob1.Text = ""
  mob2.Text = ""
  purchaseDate.Value = Now
  NxtAlert.Text = ""
End Sub

Private Sub btn_Back_Click()
Form2.Hide
End Sub

Private Sub btn_ClrAll_Click()
Call clrContents
End Sub


Private Sub btn_Updt_Click()
  Dim strength, NewuserCount As Integer
  Dim recLen As String * 6
  Dim NextAlertDays As Integer
      
    
  strength = Val(chkSignalStrength)
 
  If CustName.Text <> "" Then 'And (mob1.Text <> "" Or mob2.Text <> "") And NxtAlert.Text <> "" Then
    NextAlertDays = Val(NxtAlert.Text) * 30
    dbrecordset.fields(0) = recordNo
    dbrecordset.fields(1) = CustName.Text
    dbrecordset.fields(2) = billDetail.Text
    dbrecordset.fields(3) = add1.Text
    dbrecordset.fields(4) = add2.Text
    dbrecordset.fields(5) = mob1.Text
    dbrecordset.fields(6) = mob2.Text
    dbrecordset.fields(7) = purchaseDate.Value
    
    dbrecordset.fields(8) = purchaseDate.Value + NextAlertDays
    dbrecordset.fields(9) = Date
    dbrecordset.fields(10) = Time
    dbrecordset.fields(11) = NxtAlert.Text
    
   
    If strength >= 9 And strength <> 99 Then
 
      Call WelcomeNewCustomer(dbrecordset.fields(5), dbrecordset.fields(6), _
            dbrecordset.fields(1), dbrecordset.fields(2), dbrecordset.fields(3), _
            dbrecordset.fields(4), dbrecordset.fields(7), dbrecordset.fields(0))
 
'      Call WelcomeNewCustomer(mob1.Text, mob2.Text, CustName.Text, _
 '            billDetail.Text, add1.Text, add2.Text, purchaseDate.Value, recordNo)
             
    Else
      Open "newUsers.txt" For Random Access Read Write As #2 Len = Len(recLen)
      
      Get #2, 1, recLen
      NewuserCount = Val(recLen)
      NewuserCount = NewuserCount + 1
      recLen = Str(NewuserCount)
      Put #2, 1, recLen
    
      recLen = Str(recordNo)
      Put #2, NewuserCount + 1, recLen
      Close #2
    End If
    
    
    dbrecordset.Update
     
     
    Call clrContents
    
    If dbrecordset.RecordCount > Form1.maxRecords - 1 Then
      MsgBox ("Memory Exhausted")
      End
    Else
      dbrecordset.AddNew
      recordNo = dbrecordset.RecordCount + 1
    End If
    purchaseDate.Value = Date
  Else
    MsgBox ("Customers Name or Mobile Number or Next Alert's Value is not metioned")
  End If
  
End Sub


Private Sub Form_Load()

  Call clrContents
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

  With Form2
    .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
  End With


  purchaseDate.Value = Date
  lbl_date.Caption = purchaseDate.Day & "-" & mmByName(Val(purchaseDate.Month)) & "-" & purchaseDate.Year

  Set db = DBEngine.Workspaces(0).OpenDatabase("db")
  Set dbrecordset = db.OpenRecordset("CUSTOMERS", dbOpenTable)

  If dbrecordset.RecordCount > Form1.maxRecords Then
    MsgBox ("Memory Exhausted")
  '  End
  Else
    dbrecordset.AddNew
    recordNo = dbrecordset.RecordCount + 1    'increment the record number
  End If
End Sub



Private Sub mnuAbout_Click()
Call aboutInfo
End Sub

Private Sub mnuLicense_Click()
Call activateLic
End Sub

Private Sub mnuRegister_Click()
If RegisterProduct = 1 Then
  MsgBox ("Product is already registered")
End If
End Sub

Private Sub mnuViewRecord_Click()

Form2.Enabled = False
Form1.ViewSource = 2
Form5.Show
End Sub


Private Sub mob1_KeyPress(KeyAscii As Integer)
  If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub mob2_KeyPress(KeyAscii As Integer)
  If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub purchaseDate_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
  lbl_date.Caption = purchaseDate.Day & "-" & mmByName(Val(purchaseDate.Month)) & "-" & purchaseDate.Year
End Sub
