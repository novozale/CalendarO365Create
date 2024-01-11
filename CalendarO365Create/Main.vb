Imports System.Net
Imports System.Net.Mail
Imports System.Text
Imports ADODB
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module Main
    Public i As Integer = 1

    Sub Main()
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Main - она и в Африке Main
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyProjectID As String
        Dim MyCompanyID As String
        Dim MyProjectName As String
        Dim MyProjectComment As String
        Dim MyStartDate As DateTime
        Dim MyFullName As String
        Dim MyUserID As Integer
        Dim MyEmail As String
        Dim MyEventID As String

        '---------------список проектов, по которым надо создать события в CRM---------------------
        MySQLStr = "SELECT tbl_CRM_Projects.ProjectID, tbl_CRM_Projects.CompanyID, "
        MySQLStr = MySQLStr + "tbl_CRM_Projects.ProjectName, tbl_CRM_Projects.ProjectComment, "
        MySQLStr = MySQLStr + "tbl_CRM_Projects.StartDate, ScalaSystemDB.dbo.ScaUsers.FullName, "
        MySQLStr = MySQLStr + "ScalaSystemDB.dbo.ScaUsers.UserID, ISNULL(RM.dbo.RM660100.RM66003, '') AS Email "
        MySQLStr = MySQLStr + "From tbl_CRM_Projects INNER JOIN "
        MySQLStr = MySQLStr + "ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Projects.ResponciblePerson = "
        MySQLStr = MySQLStr + "ScalaSystemDB.dbo.ScaUsers.FullName LEFT OUTER JOIN "
        MySQLStr = MySQLStr + "RM.dbo.RM660100 ON tbl_CRM_Projects.ResponciblePerson = RM.dbo.RM660100.RM66002 LEFT OUTER JOIN "
        MySQLStr = MySQLStr + "(SELECT ProjectID "
        MySQLStr = MySQLStr + "From tbl_CRM_Events "
        MySQLStr = MySQLStr + "WHERE (ActionTime > GETDATE()) AND (ActionDescription = N'Обновление информации по проекту') "
        MySQLStr = MySQLStr + "GROUP BY ProjectID "
        MySQLStr = MySQLStr + "HAVING (ProjectID IS NOT NULL)) AS View_8 ON tbl_CRM_Projects.ProjectID = View_8.ProjectID "
        MySQLStr = MySQLStr + "WHERE (tbl_CRM_Projects.CloseDate IS NULL) "
        MySQLStr = MySQLStr + "And (tbl_CRM_Projects.ResponciblePerson <> N'') "
        MySQLStr = MySQLStr + "And (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) "
        MySQLStr = MySQLStr + "And (View_8.ProjectID Is NULL) "
        MySQLStr = MySQLStr + "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName, tbl_CRM_Projects.StartDate "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            Do While Declarations.MyRec.EOF = False
                MyProjectID = Declarations.MyRec.Fields("ProjectID").Value
                MyCompanyID = Declarations.MyRec.Fields("CompanyID").Value
                MyProjectName = Declarations.MyRec.Fields("ProjectName").Value
                MyProjectComment = Declarations.MyRec.Fields("ProjectComment").Value
                MyStartDate = Declarations.MyRec.Fields("StartDate").Value
                MyFullName = Declarations.MyRec.Fields("FullName").Value
                MyUserID = Declarations.MyRec.Fields("UserID").Value
                MyEmail = Declarations.MyRec.Fields("Email").Value
                If MyEmail.Equals("") Then
                    SendEmailAbsence(MyFullName)
                End If
                MyEventID = CreateEvent(MyProjectID, MyCompanyID, MyProjectName, MyProjectComment,
                MyStartDate, MyFullName, MyUserID, MyEmail)
                If MyEventID.Equals("") Then
                    SendEmailError(MyProjectID, MyCompanyID, MyProjectName, MyProjectComment,
                                      MyStartDate, MyFullName, MyUserID, MyEmail)
                Else
                    If CreateCalendarEvent(MyProjectID, MyCompanyID, MyProjectName, MyProjectComment,
                                      MyStartDate, MyFullName, MyUserID, MyEmail, MyEventID) = True Then
                        SendEmailInfo(MyProjectID, MyCompanyID, MyProjectName, MyProjectComment,
                        MyStartDate, MyFullName, MyUserID, MyEmail, MyEventID)
                    Else
                        SendCalendarInfo(MyProjectID, MyCompanyID, MyProjectName, MyProjectComment,
                        MyStartDate, MyFullName, MyUserID, MyEmail, MyEventID)
                    End If
                End If
                '-----Первый запуск
                i = i + 1
                If i > 30 Then
                    i = 1
                End If
                '-----Первый запуск
                Console.WriteLine("-----> Проект " + MyProjectName + " Клиент " + MyCompanyID + " Продавец " + MyFullName + " " + MyEmail)
                Declarations.MyRec.MoveNext()
            Loop
            trycloseMyRec()
        End If
    End Sub

    Private Function CreateEvent(MyProjectID As String, MyCompanyID As String, MyProjectName As String,
                                        MyProjectComment As String, MyStartDate As DateTime, MyFullName As String,
                                        MyUserID As Integer, MyEmail As String) As String
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// создание события в CRM
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim UConn As ADODB.Connection                     'соединение с БД
        Dim URec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет
        Dim MySQLStr As String
        Dim MyGUID As Guid
        Dim MyDate As DateTime
        Dim MyContactGUID As Guid
        Dim MyContactID As String

        MyContactID = ""
        Try
            MyGUID = Guid.NewGuid
            '-----первый запуск
            MyDate = DateAdd(DateInterval.Day, i, Now())
            '-----постоянная работа
            'MyDate = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, Now()))
            If Weekday(MyDate, FirstDayOfWeek.Monday) = 6 Then
                MyDate = DateAdd(DateInterval.Day, -1, MyDate)
            End If
            If Weekday(MyDate, FirstDayOfWeek.Monday) = 7 Then
                MyDate = DateAdd(DateInterval.Day, 1, MyDate)
            End If

            MySQLStr = "SELECT TOP 1 ContactID "
            MySQLStr = MySQLStr + "From tbl_CRM_Contacts "
            MySQLStr = MySQLStr + "WHERE (CompanyID = '" + MyCompanyID + "') "
            UConn = New ADODB.Connection
            UConn.CursorLocation = 3
            UConn.CommandTimeout = 600
            UConn.ConnectionTimeout = 300
            UConn.Open(Declarations.MyConnStr)
            URec = New ADODB.Recordset
            URec.LockType = LockTypeEnum.adLockOptimistic
            URec.Open(MySQLStr, UConn)
            If URec.EOF() = True And URec.BOF = True Then
                URec.Close()
                '---Создание контакта
                MyContactGUID = Guid.NewGuid
                MyContactID = MyContactGUID.ToString
                MySQLStr = "INSERT INTO tbl_CRM_Contacts "
                MySQLStr = MySQLStr + "(ContactID, CompanyID, ContactName, ContactPhone, "
                MySQLStr = MySQLStr + "ContactEMail, FromScala, CreationDate, Comments) "
                MySQLStr = MySQLStr + "VALUES ("
                MySQLStr = MySQLStr + "'" + MyContactID + "', "
                MySQLStr = MySQLStr + "'" + MyCompanyID + "', "
                MySQLStr = MySQLStr + "N'Контактное лицо по проекту', "
                MySQLStr = MySQLStr + "N'', "
                MySQLStr = MySQLStr + "N'', "
                MySQLStr = MySQLStr + "0, "
                MySQLStr = MySQLStr + "GETDATE(), "
                MySQLStr = MySQLStr + "N'') "
                UConn.Execute(MySQLStr)
            Else
                URec.MoveFirst()
                Do While URec.EOF = False
                    MyContactID = URec.Fields("ContactID").Value
                    URec.MoveNext()
                Loop
                URec.Close()
            End If



            MySQLStr = "INSERT INTO tbl_CRM_Events "
            MySQLStr = MySQLStr + "(EventID, DirectionID, EventTypeID, EventTypeDescription, CompanyID, "
            MySQLStr = MySQLStr + "ContactID, ActionTime, ActionID, ActionDescription, ActionPlannedDate, "
            MySQLStr = MySQLStr + "ActionSumm, ActionComments, ActionResultID, ActionResultDescription, "
            MySQLStr = MySQLStr + "UserID, OwnerID, ActionClosed, ProjectID, TransportID, TransportDistance, IsApproved) "
            MySQLStr = MySQLStr + "VALUES ("
            MySQLStr = MySQLStr + "'" + MyGUID.ToString + "', "                                         '--EventID
            MySQLStr = MySQLStr + "2, "
            MySQLStr = MySQLStr + "1, "
            MySQLStr = MySQLStr + "NULL, "
            MySQLStr = MySQLStr + "'" + MyCompanyID + "', "                                             '--CompanyID
            MySQLStr = MySQLStr + "'" + MyContactID + "', "                                             '--ContactID
            MySQLStr = MySQLStr + "GETDATE(), "                                                         '--ActionTime
            MySQLStr = MySQLStr + "999999, "
            MySQLStr = MySQLStr + "N'Обновление информации по проекту', "
            MySQLStr = MySQLStr + "CONVERT(DATETIME, '" + Format(MyDate, "dd/MM/yyyy") + "', 103), "    '--ActionPlannedDate
            MySQLStr = MySQLStr + "1, "
            MySQLStr = MySQLStr + "N'', "
            MySQLStr = MySQLStr + "NULL, "
            MySQLStr = MySQLStr + "NULL, "
            MySQLStr = MySQLStr + MyUserID.ToString + ", "                                              '--UserID
            MySQLStr = MySQLStr + MyUserID.ToString + ", "                                              '--OwnerID
            MySQLStr = MySQLStr + "NULL, "
            MySQLStr = MySQLStr + "'" + MyProjectID + "', "                                             '--ProjectID
            MySQLStr = MySQLStr + "NULL, "
            MySQLStr = MySQLStr + "NULL, "
            MySQLStr = MySQLStr + "1 "
            MySQLStr = MySQLStr + ") "
            UConn.Execute(MySQLStr)

            CreateEvent = MyGUID.ToString
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Information, "Внимание!")
            EventLog.WriteEntry("CalendarO365Create", "CreateEvent --1--> " & ex.Message)
            CreateEvent = ""
        End Try
    End Function

    Private Function CreateCalendarEvent(MyProjectID As String, MyCompanyID As String, MyProjectName As String,
                                        MyProjectComment As String, MyStartDate As DateTime, MyFullName As String,
                                        MyUserID As Integer, MyEmail As String, MyEventID As String) As Boolean
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// создание события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim UConn As ADODB.Connection                     'соединение с БД
        Dim URec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет
        Dim MySQLStr As String
        Dim MyCompanyCode As String
        Dim MyCompanyName As String
        Dim MySchDate As DateTime
        Dim MyBody As String
        Dim MyToken As String

        'MyEmail = "alexander.novozhilov@elektroskandia.ru"

        Try
            MyCompanyCode = ""
            MyCompanyName = ""
            MySQLStr = "SELECT tbl_CRM_Companies.ScalaCustomerCode, tbl_CRM_Companies.CompanyName, tbl_CRM_Events.ActionPlannedDate "
            MySQLStr = MySQLStr + "From tbl_CRM_Events INNER JOIN "
            MySQLStr = MySQLStr + "tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID "
            MySQLStr = MySQLStr + "WHERE (tbl_CRM_Events.EventID = '" + MyEventID + "')"
            UConn = New ADODB.Connection
            UConn.CursorLocation = 3
            UConn.CommandTimeout = 600
            UConn.ConnectionTimeout = 300
            UConn.Open(Declarations.MyConnStr)
            URec = New ADODB.Recordset
            URec.LockType = LockTypeEnum.adLockOptimistic
            URec.Open(MySQLStr, UConn)
            If URec.EOF() = True And URec.BOF = True Then
                URec.Close()
            Else
                URec.MoveFirst()
                Do While URec.EOF = False
                    MyCompanyCode = URec.Fields("ScalaCustomerCode").Value
                    MyCompanyName = URec.Fields("CompanyName").Value
                    MySchDate = URec.Fields("ActionPlannedDate").Value
                    URec.MoveNext()
                Loop
                URec.Close()
            End If

            MyBody = "Уточнить и при необходимости изменить в CRM статус следующего проекта:" + "\n" + "\n"
            MyBody = MyBody + "Клиент:     " + MyCompanyCode + "    " + Replace(MyCompanyName, """", "'") + "\n"
            MyBody = MyBody + "Код проекта:     " + MyProjectID.ToString + "\n"
            MyBody = MyBody + "Название проекта:        " + Replace(MyProjectName, """", "'") + "\n"
            MyBody = MyBody + "Дата начала проекта:     " + Format(MyStartDate, "dd/MM/yyyy") + "\n"
            MyBody = MyBody + "Комментарий по проекту:      " + Replace(MyProjectComment, """", "'")

            MyToken = getMyToken()
            If MyToken.Equals("") Then
                CreateCalendarEvent = False
                Exit Function
            End If

            If CreateCalendarEventHTTP(MyToken, MyEmail, MyBody, MySchDate) = True Then
                CreateCalendarEvent = True
            Else
                CreateCalendarEvent = False
            End If

        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Information, "Внимание!")
            EventLog.WriteEntry("CalendarO365Create", "CreateCalendarEvent --1--> " & ex.Message)
            CreateCalendarEvent = False
        End Try
    End Function

    Private Function CreateCalendarEventHTTP(MyToken As String, MyEmail As String, MyBody As String,
                                             MySchDate As DateTime) As Boolean
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// создание записи в календарь при помощи WebClient
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim client As WebClient = New WebClient()
        Dim MyURI As String
        Dim MyBodyStr As String
        Dim MyBodyArr As Byte()
        Dim MyRespStr As String
        Dim MyRespArr As Byte()
        Dim MyDate As String
        Dim MyGUID As Guid

        Try
            MyGUID = Guid.NewGuid
            MyDate = Format(MySchDate, "yyyy-MM-dd")
            MyURI = "https://graph.microsoft.com/v1.0/users/" + MyEmail + "/calendar/events"
            MyBodyStr = "{" + vbCrLf
            MyBodyStr = MyBodyStr + "   ""subject"": ""Задача на обновление состояния проекта""," + vbCrLf
            MyBodyStr = MyBodyStr + "   ""body"" :  {" + vbCrLf
            MyBodyStr = MyBodyStr + "       ""contentType"": ""text""," + vbCrLf
            MyBodyStr = MyBodyStr + "       ""content"": """ + MyBody + """" + vbCrLf
            MyBodyStr = MyBodyStr + "   }," + vbCrLf
            MyBodyStr = MyBodyStr + "   ""start"": {" + vbCrLf
            MyBodyStr = MyBodyStr + "       ""dateTime"": """ + MyDate + "T09:00:00""," + vbCrLf
            MyBodyStr = MyBodyStr + "       ""timeZone"": ""Russian Standard Time""" + vbCrLf
            MyBodyStr = MyBodyStr + "   }," + vbCrLf
            MyBodyStr = MyBodyStr + "   ""end"": {" + vbCrLf
            MyBodyStr = MyBodyStr + "       ""dateTime"": """ + MyDate + "T17:30:00""," + vbCrLf
            MyBodyStr = MyBodyStr + "       ""timeZone"": ""Russian Standard Time""" + vbCrLf
            MyBodyStr = MyBodyStr + "   }," + vbCrLf
            MyBodyStr = MyBodyStr + "   ""transactionId"":""" + MyGUID.ToString + """" + vbCrLf
            MyBodyStr = MyBodyStr + "}"

            client.Headers.Add("Content-Type", "application/json")
            client.Headers.Add("Authorization", "Bearer " + MyToken)
            MyBodyArr = Encoding.UTF8.GetBytes(MyBodyStr)

            MyRespArr = client.UploadData(MyURI, "POST", MyBodyArr)
            MyRespStr = Encoding.UTF8.GetString(MyRespArr)

            CreateCalendarEventHTTP = True
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Information, "Внимание!")
            EventLog.WriteEntry("CalendarO365Create", "CreateCalendarEventHTTP --1--> " & ex.Message)
            CreateCalendarEventHTTP = False
        End Try

    End Function

    Private Function getMyToken() As String
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение токена для работы с MS Graph
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim client As WebClient = New WebClient()
        Dim MyURI As String
        Dim MyBodyStr As String
        Dim MyBodyArr As Byte()
        Dim MyRespStr As String
        Dim MyRespArr As Byte()
        Dim json As JObject
        Dim MyToken As String

        MyToken = ""
        Try
            MyURI = "https://login.microsoftonline.com/f91cd4eb-0e4b-4bcc-982e-32c194cfcefa/oauth2/v2.0/token"
            MyBodyStr = "client_id=917b7b74-f5b2-45c5-ae05-6b4c895b4e79"
            MyBodyStr = MyBodyStr + "&scope=https://graph.microsoft.com/.default"
            MyBodyStr = MyBodyStr + "&client_secret=K3~MGDdZH553O61c-8_m.83S682tzUvg-2"
            MyBodyStr = MyBodyStr + "&grant_type=client_credentials"

            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded")
            MyBodyArr = Encoding.UTF8.GetBytes(MyBodyStr)

            MyRespArr = client.UploadData(MyURI, "POST", MyBodyArr)
            MyRespStr = Encoding.UTF8.GetString(MyRespArr)

            json = JObject.Parse(MyRespStr)
            MyToken = json.SelectToken("access_token").ToString

        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Information, "Внимание!")
            EventLog.WriteEntry("CalendarO365Create", "getMyToken --1--> " & ex.Message)
        End Try
        getMyToken = MyToken
    End Function

    Private Sub SendCalendarInfo(MyProjectID As String, MyCompanyID As String, MyProjectName As String,
                                        MyProjectComment As String, MyStartDate As DateTime, MyFullName As String,
                                        MyUserID As Integer, MyEmail As String, MyEventID As String)
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// отправка письма с созданием события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim UConn As ADODB.Connection                     'соединение с БД
        Dim URec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет
        Dim MySQLStr As String
        Dim mMailMessage As New MailMessage()
        Dim MyEmailFrom As String
        Dim MyEmailTo As String
        Dim MySubject As String
        Dim MySmtp As String
        Dim MyPort As Integer
        Dim MyDate As DateTime
        Dim MyCompanyCode As String
        Dim MyCompanyName As String
        Dim MySchDate As DateTime
        Dim MyGUID As Guid

        Try
            MyCompanyCode = ""
            MyCompanyName = ""
            MySQLStr = "SELECT tbl_CRM_Companies.ScalaCustomerCode, tbl_CRM_Companies.CompanyName, tbl_CRM_Events.ActionPlannedDate "
            MySQLStr = MySQLStr + "From tbl_CRM_Events INNER JOIN "
            MySQLStr = MySQLStr + "tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID "
            MySQLStr = MySQLStr + "WHERE (tbl_CRM_Events.EventID = '" + MyEventID + "')"
            UConn = New ADODB.Connection
            UConn.CursorLocation = 3
            UConn.CommandTimeout = 600
            UConn.ConnectionTimeout = 300
            UConn.Open(Declarations.MyConnStr)
            URec = New ADODB.Recordset
            URec.LockType = LockTypeEnum.adLockOptimistic
            URec.Open(MySQLStr, UConn)
            If URec.EOF() = True And URec.BOF = True Then
                URec.Close()
            Else
                URec.MoveFirst()
                Do While URec.EOF = False
                    MyCompanyCode = URec.Fields("ScalaCustomerCode").Value
                    MyCompanyName = URec.Fields("CompanyName").Value
                    MySchDate = URec.Fields("ActionPlannedDate").Value
                    URec.MoveNext()
                Loop
                URec.Close()
            End If
            MyGUID = Guid.NewGuid

            MyEmailFrom = "reportserver@elektroskandia.ru"
            MyEmailTo = MyEmail
            'MyEmailTo = "cio@elektroskandia.ru"
            MySubject = "Задача на обновление состояния проекта"
            MySmtp = "spbprd5.eskru.local"
            MyPort = 25

            '-----первый запуск
            MyDate = DateAdd(DateInterval.Day, i, Now())
            '-----постоянная работа
            'MyDate = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, Now()))

            If Weekday(MyDate, FirstDayOfWeek.Monday) = 6 Then
                MyDate = DateAdd(DateInterval.Day, -1, MyDate)
            End If
            If Weekday(MyDate, FirstDayOfWeek.Monday) = 7 Then
                MyDate = DateAdd(DateInterval.Day, 1, MyDate)
            End If

            Dim str As StringBuilder
            str = New StringBuilder
            str.AppendLine("BEGIN:VCALENDAR")
            str.AppendLine("PRODID:-//Elektroskandia Rus")
            str.AppendLine("VERSION:2.0")
            str.AppendLine("METHOD:REQUEST")
            str.AppendLine("X-MS-OLK-FORCEINSPECTOROPEN:TRUE")
            str.AppendLine("BEGIN:VEVENT")
            str.AppendLine("ATTENDEE;RSVP=TRUE;ROLE=REQ-PARTICIPANT;CUTYPE=GROUP:MAILTO:" & MyEmailTo)
            str.AppendLine("STATUS:CONFIRMED")
            str.AppendLine("DTSTART:" & Format(MyDate, "yyyyMMdd") & "T050000Z")
            str.AppendLine("DTSTAMP:" & Format(MyDate, "yyyyMMdd") & "T050000Z")
            str.AppendLine("DTEND:" & Format(MyDate, "yyyyMMdd") & "T133000Z")
            str.AppendLine("SEQUENCE:0")
            str.AppendLine("LOCATION: Russia")
            str.AppendLine("UID:" & MyGUID.ToString)
            str.AppendLine("CLASS:PUBLIC")
            str.AppendLine(String.Format("DESCRIPTION:{0}", "Уточнить и при необходимости изменить в CRM статус следующего проекта: \n" _
                + "Клиент:     " + MyCompanyCode + "    " + MyCompanyName + "\n" _
                + "Код проекта:     " + MyProjectID + "\n" _
                + "Название проекта:        " + MyProjectName + "\n" _
                + "Дата начала проекта:     " + Format(MyStartDate, "dd/MM/yyyy") + "\n" _
                + "Комментарий по проекту:      " + MyProjectComment))
            str.AppendLine(String.Format("SUMMARY:{0}", "Задача на обновление состояния проекта"))
            str.AppendLine("BEGIN:VALARM")
            str.AppendLine("TRIGGER:-PT15M")
            str.AppendLine("ACTION:DISPLAY")
            str.AppendLine("DESCRIPTION:Reminder")
            str.AppendLine("END:VALARM")
            str.AppendLine("END:VEVENT")
            str.AppendLine("END:VCALENDAR")
            Dim ct As System.Net.Mime.ContentType
            ct = New System.Net.Mime.ContentType("text/calendar")
            ct.Parameters.Add("method", "REQUEST")
            ct.Parameters.Add("name", "event.ics")

            Dim avCal As AlternateView
            avCal = AlternateView.CreateAlternateViewFromString(str.ToString(), ct)
            mMailMessage.AlternateViews.Add(avCal)

            mMailMessage.From = New MailAddress(MyEmailFrom)
            mMailMessage.To.Add(New MailAddress(MyEmailTo))
            mMailMessage.Subject = MySubject
            mMailMessage.Priority = MailPriority.Normal

            ' Instantiate a new instance of SmtpClient
            Dim mSmtpClient As New SmtpClient(MySmtp, MyPort)
            mSmtpClient.Credentials = New System.Net.NetworkCredential()

            ' Send the mail message
            mSmtpClient.Send(mMailMessage)
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Information, "Внимание!")
            EventLog.WriteEntry("CalendarO365Create", "SendCalendarInfo --1--> " & ex.Message)
        End Try
    End Sub
    Private Sub SendEmailInfo(MyProjectID As String, MyCompanyID As String, MyProjectName As String,
                                        MyProjectComment As String, MyStartDate As DateTime, MyFullName As String,
                                        MyUserID As Integer, MyEmail As String, MyEventID As String)
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Отправка сообщения о том, что пользователю создана задача
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim UConn As ADODB.Connection                     'соединение с БД
        Dim URec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет
        Dim MySQLStr As String
        Dim mMailMessage As New MailMessage()
        Dim MyEmailFrom As String
        Dim MyEmailTo As String
        Dim MySubject As String
        Dim MyBody As String
        Dim MySmtp As String
        Dim MyPort As Integer
        Dim MyCompanyCode As String
        Dim MyCompanyName As String
        Dim MySchDate As DateTime

        Try
            MyCompanyCode = ""
            MyCompanyName = ""
            MySQLStr = "SELECT tbl_CRM_Companies.ScalaCustomerCode, tbl_CRM_Companies.CompanyName, tbl_CRM_Events.ActionPlannedDate "
            MySQLStr = MySQLStr + "From tbl_CRM_Events INNER JOIN "
            MySQLStr = MySQLStr + "tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID "
            MySQLStr = MySQLStr + "WHERE (tbl_CRM_Events.EventID = '" + MyEventID + "')"
            UConn = New ADODB.Connection
            UConn.CursorLocation = 3
            UConn.CommandTimeout = 600
            UConn.ConnectionTimeout = 300
            UConn.Open(Declarations.MyConnStr)
            URec = New ADODB.Recordset
            URec.LockType = LockTypeEnum.adLockOptimistic
            URec.Open(MySQLStr, UConn)
            If URec.EOF() = True And URec.BOF = True Then
                URec.Close()
            Else
                URec.MoveFirst()
                Do While URec.EOF = False
                    MyCompanyCode = URec.Fields("ScalaCustomerCode").Value
                    MyCompanyName = URec.Fields("CompanyName").Value
                    MySchDate = URec.Fields("ActionPlannedDate").Value
                    URec.MoveNext()
                Loop
                URec.Close()
            End If

            MyEmailFrom = "reportserver@elektroskandia.ru"
            MyEmailTo = MyEmail
            'MyEmailTo = "cio@elektroskandia.ru"
            MySubject = "Задача в CRM"
            MySmtp = "spbprd5.eskru.local"
            MyPort = 25
            MyBody = "Уважаемый " + MyFullName + "," + Chr(13) + Chr(10) + Chr(13) + Chr(10)
            MyBody = MyBody + "В CRM на " + Format(MySchDate, "dd/MM/yyyy") + " вам сормирована задача " + Chr(13) + Chr(10) + Chr(13) + Chr(10)
            MyBody = MyBody + "Уточнить и при необходимости изменить в CRM статус следующего проекта:" + Chr(13) + Chr(10)
            MyBody = MyBody + "Клиент:     " + MyCompanyCode + "    " + MyCompanyName + Chr(13) + Chr(10)
            MyBody = MyBody + "Код проекта:     " + MyProjectID + Chr(13) + Chr(10)
            MyBody = MyBody + "Название проекта:        " + MyProjectName + Chr(13) + Chr(10)
            MyBody = MyBody + "Дата начала проекта:     " + Format(MyStartDate, "dd/MM/yyyy") + Chr(13) + Chr(10)
            MyBody = MyBody + "Комментарий по проекту:      " + MyProjectComment + Chr(13) + Chr(10) + Chr(13) + Chr(10)


            MyBody = MyBody + "-------------------------" + Chr(13) + Chr(10)
            MyBody = MyBody + "Сервер рассылки"

            mMailMessage.From = New MailAddress(MyEmailFrom)
            mMailMessage.To.Add(New MailAddress(MyEmailTo))
            mMailMessage.Subject = MySubject
            mMailMessage.Body = MyBody
            mMailMessage.Priority = MailPriority.Normal

            ' Instantiate a new instance of SmtpClient
            Dim mSmtpClient As New SmtpClient(MySmtp, MyPort)
            mSmtpClient.Credentials = New System.Net.NetworkCredential()

            ' Send the mail message
            mSmtpClient.Send(mMailMessage)
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Information, "Внимание!")
            EventLog.WriteEntry("CalendarO365Create", "SendEmailInfo --1--> " & ex.Message)
        End Try
    End Sub

    Private Sub SendEmailError(MyProjectID As String, MyCompanyID As String, MyProjectName As String,
                                        MyProjectComment As String, MyStartDate As DateTime, MyFullName As String,
                                        MyUserID As Integer, MyEmail As String)
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Отправка сообщения о том, что не удалось сформировать задачу в CRM
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim mMailMessage As New MailMessage()
        Dim MyEmailFrom As String
        Dim MyEmailTo As String
        Dim MySubject As String
        Dim MyBody As String
        Dim MySmtp As String
        Dim MyPort As Integer

        Try
            MyEmailFrom = "reportserver@elektroskandia.ru"
            MyEmailTo = "itdep@elektroskandia.ru"
            'MyEmailTo = "cio@elektroskandia.ru"
            MySubject = "Не удалось сформировать задачу в CRM"
            MySmtp = "spbprd5.eskru.local"
            MyPort = 25
            MyBody = "Для пользователя " + MyFullName + Chr(13) + Chr(10)
            MyBody = MyBody + "В CRM не удалось автоматически сформировать следующую задачу: " + Chr(13) + Chr(10) + Chr(13) + Chr(10)
            MyBody = MyBody + "Уточнить и при необходимости изменить в CRM статус следующего проекта:" + Chr(13) + Chr(10)
            MyBody = MyBody + "Код проекта:     " + MyProjectID + Chr(13) + Chr(10)
            MyBody = MyBody + "Название проекта:        " + MyProjectName + Chr(13) + Chr(10)
            MyBody = MyBody + "Дата начала проекта:     " + Format(MyStartDate, "dd/MM/yyyy") + Chr(13) + Chr(10)
            MyBody = MyBody + "Комментарий по проекту:      " + MyProjectComment + Chr(13) + Chr(10) + Chr(13) + Chr(10)

            MyBody = MyBody + "-------------------------" + Chr(13) + Chr(10)
            MyBody = MyBody + "Сервер рассылки"

            mMailMessage.From = New MailAddress(MyEmailFrom)
            mMailMessage.To.Add(New MailAddress(MyEmailTo))
            mMailMessage.Subject = MySubject
            mMailMessage.Body = MyBody
            mMailMessage.Priority = MailPriority.Normal

            ' Instantiate a new instance of SmtpClient
            Dim mSmtpClient As New SmtpClient(MySmtp, MyPort)
            mSmtpClient.Credentials = New System.Net.NetworkCredential()

            ' Send the mail message
            mSmtpClient.Send(mMailMessage)
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Information, "Внимание!")
            EventLog.WriteEntry("CalendarO365Create", "SendEmailError --1--> " & ex.Message)
        End Try
    End Sub

    Private Sub SendEmailAbsence(MyFullName As String)
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Отправка сообщения о том, что для пользователя нет почты
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim mMailMessage As New MailMessage()
        Dim MyEmailFrom As String
        Dim MyEmailTo As String
        Dim MySubject As String
        Dim MyBody As String
        Dim MySmtp As String
        Dim MyPort As Integer

        Try
            MyEmailFrom = "reportserver@elektroskandia.ru"
            MyEmailTo = "itdep@elektroskandia.ru"
            'MyEmailTo = "cio@elektroskandia.ru"
            MySubject = "Отсутствие почты у пользователя"
            MySmtp = "spbprd5.eskru.local"
            MyPort = 25
            MyBody = "Для пользователя " + MyFullName + Chr(13) + Chr(10)
            MyBody = MyBody + "В CRM (БД RM) не указана почта." + Chr(13) + Chr(10) + Chr(13) + Chr(10)
            MyBody = MyBody + "-------------------------" + Chr(13) + Chr(10)
            MyBody = MyBody + "Сервер рассылки"

            mMailMessage.From = New MailAddress(MyEmailFrom)
            mMailMessage.To.Add(New MailAddress(MyEmailTo))
            mMailMessage.Subject = MySubject
            mMailMessage.Body = MyBody
            mMailMessage.Priority = MailPriority.Normal

            ' Instantiate a new instance of SmtpClient
            Dim mSmtpClient As New SmtpClient(MySmtp, MyPort)
            mSmtpClient.Credentials = New System.Net.NetworkCredential()

            ' Send the mail message
            mSmtpClient.Send(mMailMessage)
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Information, "Внимание!")
            EventLog.WriteEntry("CalendarO365Create", "SendEmailAbsence --1--> " & ex.Message)
        End Try
    End Sub
End Module
