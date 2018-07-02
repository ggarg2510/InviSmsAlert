Attribute VB_Name = "Module1"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public IMEIstr As String * 15
Public valuePasser5To3 As Integer 'used for passing values from form 5 to 3


Public Function chkGSMstatus()
Dim result As Byte
Dim GSM_commands(5) As String
Form1.MSComm1.InputLen = 0

GSM_commands(1) = "ATE0"
GSM_commands(2) = "AT"
GSM_commands(3) = "AT+CLIP=1"
GSM_commands(4) = "AT+CMGF=1"
GSM_commands(5) = "AT+CNMI=2,2,0,0,0"

i = 1
result = 1
While result <> 0
  Form1.InMsg = ""
'  Form1.MSComm1.Output = GSM_commands(i) + vbCrLf
  Sleep 200
  Form1.InMsg = StrConv(Form1.MSComm1.Input, vbUpperCase)
  result = InStr(Form1.InMsg, "OK" + vbCrLf)
  i = i + 1
  If i = 6 Then
    result = 0
  End If
 Wend
   If i <> 6 Then
    chkGSMstatus = 0
    Form1.StatusBar1.Panels(4).Text = "NOT CONNECTED"
  Else
    chkGSMstatus = 1
    Form1.StatusBar1.Panels(4).Text = "FOUND"
  End If
Form1.InMsg = ""
Form1.MSComm1.InputLen = 1
  Call chkSignalStrength
End Function


Public Function chkSignalStrength()
  Dim result As Byte
  Dim signal As String
  Form1.MSComm1.InputLen = 0
  chkSignalStrength = 0
  Form1.InMsg = ""
  Form1.MSComm1.Output = "AT+CSQ" + vbCrLf
  Sleep 200
  Form1.InMsg = StrConv(Form1.MSComm1.Input, vbUpperCase)
  result = InStr(Form1.InMsg, Chr(32))
  If InStr(Form1.InMsg, vbCrLf + "OK") <> 0 Then
    Form1.InMsg = Mid(Form1.InMsg, result, InStr(Form1.InMsg, vbCrLf + "OK") - result)
    Form1.StatusBar1.Panels(5).Text = Form1.InMsg
    Form1.StatusBar1.Panels(4).Text = "CONNECTED"
  Else
    Form1.StatusBar1.Panels(5).Text = "NULL"
    Form1.StatusBar1.Panels(4).Text = "NOT CONNECTED"
  End If
  
  Form1.MSComm1.InputLen = 1
  chkSignalStrength = Form1.InMsg
  Form1.InMsg = ""
End Function

Function chkGSmConnect() As Byte
  Dim result As Byte
  chkGSmConnect = 0
  Form1.MSComm1.InputLen = 0
  Form1.InMsg = ""
  Form1.MSComm1.Output = "AT" + vbCrLf
  Sleep 200
  Form1.InMsg = StrConv(Form1.MSComm1.Input, vbUpperCase)
  result = InStr(Form1.InMsg, "OK" + vbCrLf)
  If result <> 0 Then
   chkGSmConnect = 1
  End If
End Function


Function IMEIno() As String
  Dim result As Byte
  Form1.MSComm1.InputLen = 0
  Form1.InMsg = ""
  Form1.MSComm1.Output = "AT+CGSN" + vbCrLf
  Sleep 200
  Form1.InMsg = StrConv(Form1.MSComm1.Input, vbUpperCase)
  result = InStr(Form1.InMsg, vbCrLf)
  If result <> 0 Then
    IMEIno = Mid(Form1.InMsg, 3, 15)
  Else
    IMEIno = 0
  End If
End Function

Function sendSMS(ByRef number As String, ByRef body As String)
  Form1.MSComm1.Output = "AT+CMGS=" + Chr(34) + number + Chr(34) + vbCrLf
  Form1.MSComm1.Output = body + Chr(26)
   
End Function

Public Function WelcomeNewCustomer(ByRef m1 As String, ByRef m2 As String, _
            ByRef Name As String, ByRef bill As String, _
            ByRef add1 As String, ByRef add2 As String, ByVal pDate As Date, _
            ByVal recordNo As Integer)
  
 Dim msgBody As String
 
    msgBody = "INVI-SMS ALERTS: WELCOME " + Chr(34) + Name + Chr(34)
    If bill <> "" Then
      msgBody = msgBody + ", With Invoice: " + Chr(34) + bill + Chr(34)
    End If
    If add1 <> "" Then
      msgBody = msgBody + ", R/O: " + add1
    End If
    If add2 <> "" Then
      msgBody = msgBody + "," + add2
    End If
    msgBody = msgBody + ", Purchased on: " + Format$(pDate, "mm-dd-yy")
    
    If m1 <> "" Then
      Form1.MSComm1.Output = "AT+CMGS=" + Chr(34) + m1 + Chr(34) + vbCrLf
      Form1.MSComm1.Output = msgBody + Chr(26)
    End If
    If m2 <> "" Then
      Form1.MSComm1.Output = "AT+CMGS=" + Chr(34) + m2 + Chr(34) + vbCrLf
      Form1.MSComm1.Output = msgBody + Chr(26)
    End If
End Function
Public Function clrPendingNewUsers() As Byte
  Dim recLen As String * 6
  Dim NewuserCount As Integer
  Dim dbrecordset As Recordset
  Dim db As Database
  Dim i As Integer
  
  Open "newUsers.txt" For Random Access Read Write As #2 Len = Len(recLen)
  
  Get #2, 1, recLen
  NewuserCount = Val(recLen)
  clrPendingNewUsers = NewuserCount
  
  If NewuserCount <> 0 Then
    Get #2, NewuserCount + 1, recLen
    
    Set db = DBEngine.Workspaces(0).OpenDatabase("db")
    Set dbrecordset = db.OpenRecordset("CUSTOMERS", dbOpenTable)

    dbrecordset.MoveFirst
    For i = 1 To dbrecordset.RecordCount
      If dbrecordset.fields(0) = Val(recLen) Then
        Call WelcomeNewCustomer(dbrecordset.fields(5), dbrecordset.fields(6), _
            dbrecordset.fields(1), dbrecordset.fields(2), dbrecordset.fields(3), _
            dbrecordset.fields(4), dbrecordset.fields(7), dbrecordset.fields(0))
            
        NewuserCount = NewuserCount - 1
        recLen = Str(NewuserCount)
        Put #2, 1, recLen
        Exit For
      Else
          dbrecordset.MoveNext
          If dbrecordset.EOF Then
              dbrecordset.MovePrevious
          End If
      End If
    Next i
  End If
  Close #2
  
End Function


Function setLicRegistration(SerialKey As String) As Byte

    setLicRegistration = 1
    If SerialKey = "SKBA15ACT" Then      '20   records  initial activation
        maxRecords = 20
    ElseIf SerialKey = "SKAB15BAS" Then  '100  records  basic edition
        maxRecords = 100
    ElseIf SerialKey = "SKAC15STD" Then  '1000 records  standard mode
        maxRecords = 1000
    ElseIf SerialKey = "SKEC15ULT" Then  '5000 records  ultimate version
        maxRecords = 5000
    Else
        setLicRegistration = 0
    End If
End Function

Function SendRegistrationSMS()
  Dim temp1 As String * 15
  Dim mobNo As String
   
    If chkGSmConnect = 1 Then
      If chkGSMstatus = 1 Then
        temp1 = IMEIno
        mobNo = InputBox("Enter Mobile Number")
        If mobNo <> "" Then
          Call sendSMS("7530848399", "REGISTER ME " + temp1 + " @" + mobNo)
          Put #1, 1, temp1
        Else
          End
        End If
      End If
    Else
      MsgBox ("GSM is either not connected or not Powered" + vbCrLf + "Connect GSM HARDWARE and try again")
      End
    End If
End Function

Function RegisterProduct() As Byte

  Dim temp1 As String * 15    'for first record
  Dim temp2 As String * 15    'for second record
  Dim temp3 As String * 15    'for third(allowable max records)
  
  'initiate function with assigning resultant to 0
  RegisterProduct = 0
  
  'open/(make new file) naming "file.txt"
  'for length equals to the length of each record
  Open "file.txt" For Random Access Read Write As #1 Len = Len(temp1)
  
  
  Get #1, 1, temp1  'obtain the first record

  If Val(temp1) = 0 Then
    'if first record is missing then get to this line
    result = MsgBox("Please Wait...." + vbCrLf + "System is being processing for one time registry mode" + vbCrLf + "It will Take 20 seconds maximum", vbQuestion)
    Sleep 5000
    Call SendRegistrationSMS
  Else
    'if the first record found then chk for the second record
    Get #1, 2, temp2
    If temp1 = temp2 Then
      RegisterProduct = 1
      
      'retrieve the serialkey entered to find the max allowed records
      Get #1, 3, temp3
      If InStr(temp3, "BA1-5ACT") <> 0 Then
        Form1.maxRecords = 20
      ElseIf InStr(temp3, "AB1-5BAS") <> 0 Then
        Form1.maxRecords = 100
      ElseIf InStr(temp3, "AC1-5STD") <> 0 Then
        Form1.maxRecords = 1000
      ElseIf InStr(temp3, "EC1-5ULT") <> 0 Then
        Form1.maxRecords = 5000
      End If
    End If
  End If
  Close #1
End Function

Function activateLic() As Byte

  Dim SerialKey, temp1 As String * 15
  Dim templong1 As Long
  Dim temp As String
  Dim temp3 As String * 15
  
  Open "file.txt" For Random Access Read Write As #1 Len = Len(temp1)
  Get #1, 1, temp1
  If Val(temp1) = 0 Then
   MsgBox ("Register Your Product First")
  Else
    templong1 = Val(Mid(temp1, 11, 5))
    temp = Str(((templong1 * 529) + 2510) Mod 1000000)
    SerialKey = InputBox("Enter Serial key...!!!")

    If SerialKey <> "" Then
      activateLic = 1
      If SerialKey = "SK-" + temp + "-BA1-5ACT" Then
        Form1.maxRecords = 20
        temp3 = "-BA1-5ACT"
      ElseIf SerialKey = "SK-" + temp + "-AB1-5BAS" Then
        Form1.maxRecords = 100
        temp3 = "-AB1-5BAS"
      ElseIf SerialKey = "SK-" + temp + "-AC1-5STD" Then
        Form1.maxRecords = 1000
        temp3 = "-AC1-5STD"
      ElseIf SerialKey = "SK-" + temp + "-EC1-5ULT" Then
        Form1.maxRecords = 5000
        temp3 = "-EC1-5ULT"
      Else
        activateLic = 0
        MsgBox ("Entered Key is Invalid...!!")
      End If
       
      If activateLic = 1 Then
        Put #1, 2, temp1
        Put #1, 3, temp3
        MsgBox ("Congratulations...!!!" + (Chr(13) + Chr(10)) + Str(Form1.maxRecords) + " Records has been unlocked sucessfully")
      End If
    End If
  End If
  Close #1
End Function

Sub aboutInfo()

  If IsProdRegistered = 1 Then
    MsgBox ("Max Records Allowed: " + Str(Form1.maxRecords))
  Else
    MsgBox ("Your Product is not registered")
  End If
End Sub

Function IsProdRegistered() As Byte
  Dim temp1 As String * 15
  Dim temp2 As String * 15
  
  IsProdRegistered = 0
  Open "file.txt" For Random Access Read Write As #1 Len = Len(temp1)
  Get #1, 1, temp1
  Get #1, 2, temp2
  If temp1 = temp2 Then
    IsProdRegistered = 1
  End If
  Close #1
End Function

Public Function todaysSMSalertsCount() As Integer

  Static tempDate As Date
  Dim todaysSMScount As Integer
  Dim recLen As String * 6
  Dim dbrecordset As Recordset
  Dim db As Database
  
  todaysSMScount = 0
  todaysSMSalertsCount = 0
  
  Set db = DBEngine.Workspaces(0).OpenDatabase("db")
  Set dbrecordset = db.OpenRecordset("CUSTOMERS", dbOpenTable)
      
  If dbrecordset.RecordCount <> 0 Then
  
    Open "alertList.txt" For Random Access Read Write As #3 Len = Len(recLen)
    
    If tempDate <> Date Then
    
      tempDate = Date
      
      dbrecordset.MoveFirst
      
      For i = 1 To dbrecordset.RecordCount
      
        If dbrecordset.fields(8) = tempDate Then
          todaysSMScount = todaysSMScount + 1
          recLen = Str(i)
          Put #3, todaysSMScount + 1, recLen
      '  Else
        End If
        dbrecordset.MoveNext
        If dbrecordset.EOF Then
          dbrecordset.MovePrevious
        End If
      Next i
      
      recLen = Str(todaysSMScount)
      Put #3, 1, recLen
    End If
    
    Get #3, 1, recLen
    todaysSMSalertsCount = Val(recLen)
    Close #3
  End If
End Function

Public Function clrPendingAlertSMS() As Byte
  Dim recLen As String * 6
  Dim AlertSMSCount As Integer
  Dim dbrecordset As Recordset
  Dim db As Database
  Dim i As Integer
  
  Open "alertList.txt" For Random Access Read Write As #3 Len = Len(recLen)
  
  Get #3, 1, recLen
  AlertSMSCount = Val(recLen)
  clrPendingAlertSMS = AlertSMSCount
  
  If AlertSMSCount <> 0 Then
    Get #3, AlertSMSCount + 1, recLen
    
    Set db = DBEngine.Workspaces(0).OpenDatabase("db")
    Set dbrecordset = db.OpenRecordset("CUSTOMERS", dbOpenTable)

    dbrecordset.MoveFirst
    For i = 1 To dbrecordset.RecordCount
      If dbrecordset.fields(0) = Val(recLen) Then
        Call sendAlertSMS(dbrecordset.fields(5), dbrecordset.fields(6), _
                          dbrecordset.fields(1), dbrecordset.fields(7))
          
        AlertSMSCount = AlertSMSCount - 1
        recLen = Str(AlertSMSCount)
        Put #3, 1, recLen
        Exit For
      Else
          dbrecordset.MoveNext
          If dbrecordset.EOF Then
              dbrecordset.MovePrevious
          End If
      End If
    Next i
  End If
  Close #3
  
End Function

Public Function sendAlertSMS(ByRef m1 As String, ByRef m2 As String, _
            ByRef Name As String, ByVal pDate As Date)
 
  Dim msgBody As String
  msgBody = "INVI-SMS ALERT: " + Name + ",You had purchased machine on: " + _
            Format$(pDate, "mm-dd-yy") + "This is to inform you that there is" + _
            " a need to replace your battery water, ignoring this may result to" + _
            " severe damage to the battery."
  If m1 <> "" Then
    Form1.MSComm1.Output = "AT+CMGS=" + Chr(34) + m1 + Chr(34) + vbCrLf
    Form1.MSComm1.Output = msgBody + Chr(26)
  End If
  If m2 <> "" Then
    Form1.MSComm1.Output = "AT+CMGS=" + Chr(34) + m2 + Chr(34) + vbCrLf
    Form1.MSComm1.Output = msgBody + Chr(26)
  End If
End Function
