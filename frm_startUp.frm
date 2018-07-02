VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INVI-SMS ALERT"
   ClientHeight    =   8685
   ClientLeft      =   -1995
   ClientTop       =   1335
   ClientWidth     =   15105
   FillStyle       =   5  'Downward Diagonal
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8685
   ScaleWidth      =   15105
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   8190
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Text            =   "SYSTEM"
            TextSave        =   "SYSTEM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14535
            MinWidth        =   14535
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Enabled         =   0   'False
            Text            =   "GSM STATUS"
            TextSave        =   "GSM STATUS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "WAIT...."
            TextSave        =   "WAIT...."
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4920
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      RThreshold      =   1
   End
   Begin VB.Timer tim_WelcomeImgTim 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5400
      Top             =   6000
   End
   Begin VB.CommandButton btn_GSM 
      Caption         =   "&SETTINGS"
      Enabled         =   0   'False
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
      Left            =   840
      TabIndex        =   3
      Top             =   6240
      Width           =   3255
   End
   Begin VB.CommandButton btn_Exit 
      Caption         =   "&EXIT"
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
      Left            =   840
      TabIndex        =   4
      Top             =   8040
      Width           =   3255
   End
   Begin VB.CommandButton btn_ViewRec 
      Caption         =   "&VIEW RECORDS"
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
      Left            =   840
      TabIndex        =   2
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton btn_AddNewRec 
      Caption         =   "&ADD RECORD"
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
      Left            =   840
      TabIndex        =   1
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Image Img_welcome 
      Height          =   7335
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "WELCOME  TO  INVI-SMS ALERT"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   10935
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'used in ImgTimer to do multiple tasks(GSM conrol)
'with just single timer
Dim GSMchk As Boolean

'when the controls goto the "View Records" control then
'this variable will keep track from where the control has
'been called. if ViewSource = 1 then it is called from
'startup form else from the addNewRecord form
Public ViewSource As Byte
              
'used for storing the incoming message from comm port
'This variable is used wherever there is a need to
'use comm port through out the project
Public InMsg As String

'used to store the max allowable record memory
Public maxRecords As Integer


Private Sub btn_AddNewRec_Click()
  'tim_WelcomeImgTim.Enabled = False
  

  If IsProdRegistered = 0 Then
    MsgBox ("First Register your Product")
    tim_WelcomeImgTim.Enabled = True
  Else
   'enter only when the product is registered
   Form2.Show
   ' tim_WelcomeImgTim.Enabled = False
  End If
End Sub

Private Sub btn_Exit_Click()
Close #1
tim_WelcomeImgTim.Enabled = False
End
End Sub

Private Sub btn_GSM_Click()
tim_WelcomeImgTim.Enabled = False
'Form4.Show
tim_WelcomeImgTim.Enabled = True
End Sub

Private Sub btn_ViewRec_Click()

  tim_WelcomeImgTim.Enabled = False
  If IsProdRegistered = 0 Then
    MsgBox ("First Register your Product")
    tim_WelcomeImgTim.Enabled = True
  Else
   'enter only when the product is registered
    Form1.Enabled = False
    ViewSource = 1
    Form5.Show
  End If
End Sub

Private Sub Form_Load()
  'adjusting the form settings like its width and height
  'we have make form width and height same as screen's
  'width and height respectively
  
  With Form1
    .Width = Screen.Width
    .Height = Screen.Height
    .Move 0, 0
  End With
  
'  MSComm1.PortOpen = True     'open the comm port
    
    
  'go for the product registration here
  'this has two primary jobs:-
  ' a)it sends SMS containing IMEI code of hardware and,
  ' b)if product is registered already, then check for the
  '   maximum available records
  Call RegisterProduct
  
  
  'enables the timer used for switching between the pictures
  tim_WelcomeImgTim.Enabled = True
  
  'tim_GSM.Enabled = True     remove it
  'tim_GSM.Interval = 2000    remove it


  'bydefault load the "2.jpg" picture
  Img_welcome.Picture = LoadPicture("ICONS_AND_IMAGES\2.jpg")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #1
End Sub

Private Sub mnuAbout_Click()
Call aboutInfo
End Sub

Private Sub mnuLicense_Click()
Call activateLic
End Sub

Private Sub mnuRegister_Click()

tim_WelcomeImgTim.Enabled = False
If RegisterProduct = 1 Then
  MsgBox ("Product is already registered")
End If
tim_WelcomeImgTim.Enabled = True
End Sub

Private Sub MSComm1_OnComm()
  Static InStatus As Byte
  Static i, j As Byte
  Static SMSMobNo As String
  
  Select Case MSComm1.CommEvent
     Case comEvReceive
      InMsg = InMsg + MSComm1.Input
  End Select




  If InStatus = 0 Then
    If InStr(InMsg, "RING" + vbCrLf + vbCrLf + "+CLIP:") <> 0 Then
      InStatus = 1          ' CALL
      MSComm1.Output = "ATH" + vbCrLf
      InMsg = ""
    ElseIf InStr(InMsg, vbCrLf + "+CMT: ") <> 0 Then
      InStatus = 2          ' SMS
      InMsg = ""
    End If
  End If
  
  
  
  
  Select Case InStatus
    Case 1                  ' CALL
      i = InStr(InMsg, Chr(43))
      If i <> 0 Then
        If InStr(InMsg, Chr(34) + Chr(44)) <> 0 Then
          StatusBar1.Panels(2).Text = "MISSED CALL FROM:  " + Mid(InMsg, i, 13)
          InMsg = ""
          InStatus = 0
          i = 0
        End If
      End If
      
    Case 2                  ' SMS
      If j = 0 Then
        i = InStr(InMsg, Chr(43))
        If i <> 0 Then
          If InStr(InMsg, Chr(34) + Chr(44)) <> 0 And SMSMobNo = "" Then
            SMSMobNo = Mid(InMsg, i, 13)
          ElseIf InStr(InMsg, vbCrLf) <> 0 Then
            InMsg = ""
            j = 1
          End If
        End If
      ElseIf InStr(InMsg, vbCrLf) <> 0 Then
        StatusBar1.Panels(2).Text = "FROM:  " + SMSMobNo + "  MESSAGE: " + InMsg
        InMsg = ""
        SMSMobNo = ""
        InStatus = 0
        i = 0
        j = 0
      End If
  End Select
  

End Sub


Private Sub tim_WelcomeImgTim_Timer()
Static ImgChangeCount As Byte
Static GSMfound As Boolean
Static count As Byte
Static sendSMSflag As Boolean
Dim strength As Integer

  tim_WelcomeImgTim.Enabled = False
  ImgChangeCount = ImgChangeCount + 1
  If ImgChangeCount = 1 Then
          Img_welcome.Picture = LoadPicture("ICONS_AND_IMAGES\1.jpg")
  ElseIf ImgChangeCount = 2 Then
          Img_welcome.Picture = LoadPicture("ICONS_AND_IMAGES\2.jpg")
  ElseIf ImgChangeCount = 3 Then
          Img_welcome.Picture = LoadPicture("ICONS_AND_IMAGES\3.jpg")
  Else
      ImgChangeCount = 1
  End If
  
  Call todaysSMSalertsCount
  
  
  If GSMfound = 0 Then
  'check for GSM hardware at every timer event if found not-connected
    If chkGSMstatus Then
      GSMfound = 1
      count = 0
    End If
  Else
    count = count + 1
    If count = 5 Then
      count = 0
      'check signal strength after 5 timer events
      chkSignalStrength
    End If
  End If
  
  If sendSMSflag = True Then
  
    strength = Val(chkSignalStrength)
    If strength >= 9 And strength <> 99 Then
      If clrPendingNewUsers = 0 Then
         'Call clrPendingAlertSMS
      End If
    End If
  End If
  
  If sendSMSflag = True Then
    sendSMSflag = False
  Else
    sendSMSflag = True
  End If
  
  
 
  
  tim_WelcomeImgTim.Enabled = True
End Sub
