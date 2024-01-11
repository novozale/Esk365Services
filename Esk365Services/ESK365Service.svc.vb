' NOTE: You can use the "Rename" command on the context menu to change the class name "Esk365Service" in code, svc and config file together.
' NOTE: In order to launch WCF Test Client for testing this service, please select Esk365Service.svc or Esk365Service.svc.vb at the Solution Explorer and start debugging.
Imports System.Data.SqlClient
Imports Microsoft.Graph

Public Class Esk365Service
    Implements IEsk365Service
    Dim connString As String = "Data Source=.;server=sqlcls;Initial Catalog=ScaDataDB;User ID=sa;Password=sqladmin"

    Public Sub New()
    End Sub

    Private Function IsAuthorised(MyLogin As String, MyService As String) As Boolean
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка - есть ли у данного пользователя право на использование сервиса
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim ds As New DataSet()
        Dim MyAuth As Boolean = False

        Try
            MySQLStr = "dbo.spp_Services_GetAuthInfo"
            Using MyConn As SqlConnection = New SqlConnection(connString)
                Try
                    Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandTimeout = 1800
                        cmd.Parameters.AddWithValue("@MyLogin", MyLogin)
                        cmd.Parameters.AddWithValue("@MyService", MyService)
                        Using da As New SqlDataAdapter()
                            da.SelectCommand = cmd
                            da.Fill(ds)
                            If ds.Tables(0).Rows.Count <> 0 Then
                                MyAuth = True
                            End If
                        End Using
                    End Using
                Catch ex As Exception
                    EventLog.WriteEntry("ESK365Services", "IsAuthorised --1--> " & ex.Message)
                Finally
                    MyConn.Close()
                End Try
            End Using
        Catch ex As Exception
            EventLog.WriteEntry("ESK365Services", "IsAuthorised --2--> " & ex.Message)
        End Try
        Return MyAuth
    End Function

    Public Async Function CreateCalendarEventAsync(ByVal MyEvent As CreateCalendarEventType) As Threading.Tasks.Task(Of String) Implements IEsk365Service.CreateCalendarEventAsync
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MyLogin As String
        Dim MyService As String
        Dim MyId As String

        MyLogin = MyEvent.Login
        MyService = "Esk365CreateCalendarEvent"

        If IsAuthorised(MyLogin, MyService) Then
            '------------в случае авторизации - создание события
            Try
                MyId = Await FCreateCalendarEvent(MyEvent)
                Return MyId
            Catch ex As Exception
                Return ""
            End Try
        Else
            Return ""
        End If
    End Function

    Private Async Function FCreateCalendarEvent(ByVal MyEvent As CreateCalendarEventType) As Threading.Tasks.Task(Of String)
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция создания события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MyCl As MSGraphClient

        MyCl = New MSGraphClient
        If MyEvent.CalendarEventIDOld.Equals("") = False Then
            '-----Удаление старой записи с ID
            Try
                Await MyCl.graphClient.Users(MyEvent.Email).Events(MyEvent.CalendarEventIDOld).Request().DeleteAsync()
            Catch ex As Exception
            End Try
        End If
        Dim MyOEvent As New Microsoft.Graph.Event
        MyOEvent.Subject = MyEvent.Subject
        MyOEvent.Body = New ItemBody
        MyOEvent.Body.Content = MyEvent.Body
        MyOEvent.Start = New DateTimeTimeZone
        'MyOEvent.Start.DateTime = MyEvent.Start
        MyOEvent.Start.DateTime = Right("00" + MyEvent.Start.Month.ToString(), 2) + "." + Right("00" + MyEvent.Start.Day.ToString(), 2) + "." + MyEvent.Start.Year.ToString() + " " + Right("00" + MyEvent.Start.Hour.ToString(), 2) + ":00:00"
        MyOEvent.Start.TimeZone = MyEvent.Timezone
        MyOEvent.End = New DateTimeTimeZone
        'MyOEvent.End.DateTime = MyEvent.Finish
        MyOEvent.End.DateTime = Right("00" + MyEvent.Finish.Month.ToString(), 2) + "." + Right("00" + MyEvent.Finish.Day.ToString(), 2) + "." + MyEvent.Finish.Year.ToString() + " " + Right("00" + MyEvent.Finish.Hour.ToString(), 2) + ":00:00"
        MyOEvent.End.TimeZone = MyEvent.Timezone
        MyOEvent.IsReminderOn = True
        Dim createdEvent = Await MyCl.graphClient.Users(MyEvent.Email).Calendar.Events.Request().AddAsync(MyOEvent)
        Return createdEvent.Id
    End Function

    Public Async Function DeleteCalendarEventAsync(ByVal MyEvent As DeleteCalendarEventType) As Threading.Tasks.Task(Of String) Implements IEsk365Service.DeleteCalendarEventAsync
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MyLogin As String
        Dim MyService As String
        Dim MyId As String

        MyLogin = MyEvent.Login
        MyService = "Esk365CreateCalendarEvent"

        If IsAuthorised(MyLogin, MyService) Then
            '------------в случае авторизации - удаление события
            Try
                MyId = Await FDeleteCalendarEvent(MyEvent)
                Return MyId
            Catch ex As Exception
                Return ""
            End Try
        Else
            Return ""
        End If
    End Function


    Private Async Function FDeleteCalendarEvent(ByVal MyEvent As DeleteCalendarEventType) As Threading.Tasks.Task(Of String)
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция удаления события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MyCl As MSGraphClient

        MyCl = New MSGraphClient
        If MyEvent.CalendarEventIDOld.Equals("") = False Then
            '-----Удаление старой записи с ID
            Try
                Await MyCl.graphClient.Users(MyEvent.Email).Events(MyEvent.CalendarEventIDOld).Request().DeleteAsync()
                Return "Success"
            Catch ex As Exception
                Return ""
            End Try
        End If
        Return ""
    End Function
End Class
