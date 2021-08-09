Function send_Email(to_Number As String, Total_Vol As Long, Ser_Lev As Double, Max_Avail As Long)
    Debug.Print "Sending: " & to_Number
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    With OutMail
        .To = to_Number
        .Body = "SCE " & Format(Now, "hh:mm AM/PM") & "  Q: " & Total_Vol & "  SL: " & Format(Ser_Lev, "00.0") & "  MaxA: " & Max_Avail
        
        'since many emails are sent out every 30 minutes we will delete them to keep the outbox clear
        .DeleteAfterSubmit = True
        .Send
    End With
    Set OutMail = Nothing
    Set OutApp = Nothing

End Function
