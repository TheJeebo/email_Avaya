Sub Queue_Check()
    'Checks queues every 30 minutes and sends a chirp
    On Error GoTo ErrHandle
    Debug.Print "Start: " & Now
    
    '//////////////////////Avaya Report\\\\\\\\\\\\\\\\\\\\\\\
    'Opens server and runs AHT-Agent Day Skill for given day, exports to "Temp" tab
    Dim cvsApp As Object
    Dim cvsConn As Object
    Dim cvsConn2 As ACSCN.cvsConnection
    Dim cvsSrv As Object
    Dim Rep As Object
    Dim Info As Object, Log As Object, b As Object
    
    Set cvsApp = CreateObject("ACSUP.cvsApplication")
    'these no longer work with update
    'Set cvsConn = CreateObject("ACSCN.cvsConnection")
    'Set cvsConn2 = New ACSCN.cvsConnection
    Set cvsSrv = CreateObject("ACSUPSRV.cvsServer")
    Set Rep = CreateObject("ACSREP.cvsReport")
    serverAddress = "10.1.1.1"
    UserName = "myUserName"
    PassW = "myPassWord"
    Application.DisplayAlerts = False
    
    If cvsApp.CreateServer(UserName, "", "", serverAddress, False, "ENU", cvsSrv, cvsConn) Then
    If cvsConn.Login(UserName, PassW, serverAddress, "ENU") Then
    On Error Resume Next
    cvsSrv.Reports.ACD = 1
    Set Info = cvsSrv.Reports.Reports("Integrated\Designer\Comparison Report - Andre 2.1")
    If Info Is Nothing Then
    If cvsSrv.Interactive Then
    MsgBox "The Report " & "Integrated\Designer\Comparison Report - Andre 2.1" & " was not found on ACD 1", vbCritical Or vbOKOnly, "CentreVu Supervisor"
    Else
    Set Log = CreateObject("ACSERR.cvslog")
    Log.AutoLogWrite "The Report " & "Integrated\Designer\Comparison Report - Andre 2.1" & " was not found on ACD 1"
    Set Log = Nothing
    End If
    Else
    b = cvsSrv.Reports.CreateReport(Info, Rep)
    If b Then
    
    Debug.Print Rep.SetProperty("Splits/Skills", "851;852;952;954;955;956;957;967;974;850;950;951;915;916;917;975;976")
    
    b = Rep.ExportData("", 9, 0, False, True, True)

    Sheets("Main").Columns("A:R").Clear
    Sheets("Main").Cells(1, 1).PasteSpecial
    Rep.Quit


    DoEvents
    End If
    
    Rep.Quit
    
    End If
    
    Set Info = Nothing
    End If
    End If
    
    If Not cvsSrv.Interactive Then
    cvsSrv.activetasks.Remove Rep.taskID
    cvsApp.servers.Remove cvsSrv.serverkey
    End If
    
    cvsConn.Logout
    cvsConn.Disconnect
    cvsSrv.Connected = False
    Set Log = Nothing
    Set Rep = Nothing
    Set cvsSrv = Nothing
    Set cvsConn = Nothing
    Set cvsApp = Nothing
    '\\\\\\\\\\\\\\\\\\\\\\End Avaya//////////////////////
    
    
    Debug.Print "End: " & Now
    
    'if avaya report is blank workbook closes
    If ThisWorkbook.Sheets(1).Range("A1") = "" Then GoTo BlankReport
    
justEmail:
    Dim Total_Vol As Long, Max_Avail As Long
    Dim Ser_Lev As Double
    Dim OutApp, OutMail As Object
    
    Total_Vol = WorksheetFunction.Sum(ThisWorkbook.Sheets(1).Range("I2:I30"))
    Ser_Lev = WorksheetFunction.SumProduct(ThisWorkbook.Sheets(1).Range("Q2:Q30"), ThisWorkbook.Sheets(1).Range("K2:K30")) / WorksheetFunction.Sum(ThisWorkbook.Sheets(1).Range("K2:K30"))
    Max_Avail = WorksheetFunction.Max(ThisWorkbook.Sheets(1).Range("C2:C30"))
    
    
    If Total_Vol > 10 Or Ser_Lev < 70 Or Max_Avail > 15 Then
        send_Email "9155551234@vtext.com", Total_Vol, Ser_Lev, Max_Avail
        send_Email "9155551235@tmomail.net", Total_Vol, Ser_Lev, Max_Avail
        send_Email "9155551236@tmomail.net", Total_Vol, Ser_Lev, Max_Avail
        send_Email "9155551237@tmomail.net", Total_Vol, Ser_Lev, Max_Avail
    End If
    
    ThisWorkbook.Save
    ThisWorkbook.Close
    
    Exit Sub
    
    
ErrHandle:
    cvsConn.Logout
    cvsConn.Disconnect
    cvsSrv.Connected = False
    Set Log = Nothing
    Set Rep = Nothing
    Set cvsSrv = Nothing
    Set cvsConn = Nothing
    Set cvsApp = Nothing
    
BlankReport:
    ThisWorkbook.Save
    ThisWorkbook.Close
End Sub
