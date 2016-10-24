Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Text
Imports System.Object
Imports WinSCP

Class studentParent
    Public student_id As Integer
    Public parent_id As Integer
End Class

Class user
    Public Delete As String
    Public SchoolboxUserID As String
    Public Username As String
    Public ExternalID As String
    Public Title As String
    Public FirstName As String
    Public Surname As String
    Public Role As String
    Public Campus As String
    Public Password As String
    Public AltEmail As String
    Public Year As String
    Public House As String
    Public ResidentialHouse As String
    Public EPortfolio As String
    Public HideContactDetails As String
    Public HideTimetable As String
    Public EmailAddressFromUsername As String
    Public UseExternalMailClient As String
    Public EnableWebmailTab As String
    Public Superuser As String
    Public AccountEnabled As String
    Public ChildExternalIDs As String
    Public DateOfBirth As String
    Public HomePhone As String
    Public MobilePhone As String
    Public WorkPhone As String
    Public Address As String
    Public Suburb As String
    Public Postcode As String
End Class

Class uploadServer
    Public host As String
    Public userName As String
    Public pass As String
    Public rsa As String
End Class


Class configSettings
    Public connectionString As String
    Public uploadServers As List(Of uploadServer)
    Public studentEmailDomain As String
End Class


Module Main

    Public Sub Main(ByVal sArgs() As String)


        Dim config As configSettings
        config = readConfig()

        Call user(config)
        Call timetableStructure(config)
        Call timetable(config)
        Call enrollment(config)
        Call uploadFiles(config)
    End Sub


    Sub user(config As configSettings)

        ' Students ****************
        Dim ConnectionString As String = config.connectionString
        Dim commandString As String = "
select
username,
student_number,
firstname,
surname,
birthdate,
form_name
from schoolbox_students
"
        Dim users As New List(Of user)



        Using conn As New System.Data.Odbc.OdbcConnection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New System.Data.Odbc.OdbcCommand(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As System.Data.Odbc.OdbcDataReader
            dr = command.ExecuteReader

            Dim i As Integer = 0
            While dr.Read()
                users.Add(New user)

                users.Last.Delete = ""
                users.Last.SchoolboxUserID = ""
                users.Last.Title = ""

                If Not dr.IsDBNull(5) Then
                    Try
                        If Int(Right(dr.GetValue(5), 2)) < 7 Then
                            users.Last.Campus = "Junior"
                            users.Last.Role = "Junior Students"
                        Else
                            users.Last.Campus = "Senior"
                            users.Last.Role = "Senior Students"
                        End If
                    Catch
                        users.Last.Campus = "Junior"
                        users.Last.Role = "Junior Students"
                    End Try
                End If

                users.Last.Password = ""
                users.Last.AltEmail = dr.GetValue(0) & config.studentEmailDomain
                users.Last.Year = ""
                users.Last.House = ""
                users.Last.ResidentialHouse = ""
                users.Last.EPortfolio = "Y"
                users.Last.HideContactDetails = "Y"
                users.Last.HideTimetable = "N"
                users.Last.EmailAddressFromUsername = "N"
                users.Last.UseExternalMailClient = "N"
                users.Last.EnableWebmailTab = "Y"
                users.Last.Superuser = "N"
                users.Last.AccountEnabled = "Y"
                users.Last.ChildExternalIDs = ""
                users.Last.HomePhone = ""
                users.Last.MobilePhone = ""
                users.Last.WorkPhone = ""
                users.Last.Address = ""
                users.Last.Suburb = ""
                users.Last.Postcode = ""



                If Not dr.IsDBNull(0) Then users.Last.Username = dr.GetValue(0)
                If Not dr.IsDBNull(1) Then users.Last.ExternalID = dr.GetValue(1)
                If Not dr.IsDBNull(2) Then users.Last.FirstName = """" & dr.GetValue(2) & """"
                If Not dr.IsDBNull(3) Then users.Last.Surname = """" & dr.GetValue(3) & """"
                If Not dr.IsDBNull(4) Then users.Last.DateOfBirth = ddMMYYYY_to_yyyyMMdd(dr.GetValue(4))



            End While
            conn.Close()
        End Using




        'Staff **********************
        commandString = "
select
username,
staff_number,
salutation,
firstname,
surname,
house
from schoolbox_staff
"

        Using conn As New System.Data.Odbc.OdbcConnection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New System.Data.Odbc.OdbcCommand(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As System.Data.Odbc.OdbcDataReader
            dr = command.ExecuteReader

            While dr.Read()
                users.Add(New user)

                users.Last.Delete = ""
                users.Last.SchoolboxUserID = ""
                users.Last.Title = ""
                users.Last.Role = "Staff"
                users.Last.Campus = "Senior"
                users.Last.Password = ""
                users.Last.AltEmail = ""
                users.Last.Year = ""
                users.Last.ResidentialHouse = ""
                users.Last.EPortfolio = "Y"
                users.Last.HideContactDetails = "Y"
                users.Last.HideTimetable = "N"
                users.Last.EmailAddressFromUsername = "Y"
                users.Last.UseExternalMailClient = "Y"
                users.Last.EnableWebmailTab = "N"
                users.Last.Superuser = "N"
                users.Last.AccountEnabled = "Y"
                users.Last.ChildExternalIDs = ""
                users.Last.HomePhone = ""
                users.Last.MobilePhone = ""
                users.Last.WorkPhone = ""
                users.Last.Address = ""
                users.Last.Suburb = ""
                users.Last.Postcode = ""
                users.Last.DateOfBirth = ""


                If Not dr.IsDBNull(0) Then users.Last.Username = dr.GetValue(0)
                If Not dr.IsDBNull(1) Then users.Last.ExternalID = dr.GetValue(1)
                If Not dr.IsDBNull(2) Then users.Last.Title = dr.GetValue(2)
                If Not dr.IsDBNull(3) Then users.Last.FirstName = """" & dr.GetValue(3) & """"
                If Not dr.IsDBNull(4) Then users.Last.Surname = """" & dr.GetValue(4) & """"
                If Not dr.IsDBNull(5) Then users.Last.House = dr.GetValue(5)





            End While
            conn.Close()
        End Using





        'Parent to student **********************
        commandString = "
select
student_number,
carer_number
from schoolbox_parent_student
"
        Dim studentParents As New List(Of studentParent)
        Using conn As New System.Data.Odbc.OdbcConnection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New System.Data.Odbc.OdbcCommand(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As System.Data.Odbc.OdbcDataReader
            dr = command.ExecuteReader

            Dim i As Integer = 0
            While dr.Read()
                studentParents.Add(New studentParent)
                If Not dr.IsDBNull(0) Then studentParents(i).student_id = dr.GetValue(0)
                If Not dr.IsDBNull(1) Then studentParents(i).parent_id = dr.GetValue(1)
                i += 1
            End While
            conn.Close()
        End Using






        'Parents **********************
        commandString = "
select
schoolbox_parents.email,
schoolbox_parents.carer_number,
contact.firstname,
contact.surname



from schoolbox_parents
left join carer on schoolbox_parents.carer_number = carer.carer_number
left join contact on carer.contact_id = contact.contact_id

"

        Using conn As New System.Data.Odbc.OdbcConnection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New System.Data.Odbc.OdbcCommand(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As System.Data.Odbc.OdbcDataReader
            dr = command.ExecuteReader

            Dim i As Integer = 0
            While dr.Read()
                users.Add(New user)

                users.Last.Delete = ""
                users.Last.SchoolboxUserID = ""
                users.Last.Title = ""
                users.Last.Role = "Parents"
                users.Last.Campus = "Senior"
                users.Last.Password = ""
                users.Last.AltEmail = ""
                users.Last.Year = ""
                users.Last.ResidentialHouse = ""
                users.Last.EPortfolio = "N"
                users.Last.HideContactDetails = "Y"
                users.Last.HideTimetable = "Y"
                users.Last.EmailAddressFromUsername = "N"
                users.Last.UseExternalMailClient = "Y"
                users.Last.EnableWebmailTab = "N"
                users.Last.Superuser = "N"
                users.Last.AccountEnabled = "Y"
                users.Last.HomePhone = ""
                users.Last.MobilePhone = ""
                users.Last.WorkPhone = ""
                users.Last.DateOfBirth = ""
                users.Last.Address = ""
                users.Last.Suburb = ""
                users.Last.Postcode = ""
                If Not dr.IsDBNull(0) Then users.Last.Username = Strings.Left(dr.GetValue(0), Strings.InStr(dr.GetValue(0), "@") - 1)
                If Not dr.IsDBNull(1) Then users.Last.ExternalID = dr.GetValue(1)
                If Not dr.IsDBNull(2) Then users.Last.FirstName = """" & dr.GetValue(2) & """"
                If Not dr.IsDBNull(3) Then users.Last.Surname = """" & dr.GetValue(3) & """"


                For Each a In studentParents
                    If users.Last.ExternalID = a.parent_id Then
                        If users.Last.ChildExternalIDs = "" Then
                            users.Last.ChildExternalIDs = a.student_id
                        Else
                            users.Last.ChildExternalIDs = users.Last.ChildExternalIDs & "," & a.student_id
                        End If
                    End If

                Next
                users.Last.ChildExternalIDs = """" & users.Last.ChildExternalIDs & """"




            End While

        End Using


        'Parents (Spouse) **********************
        commandString = "
select
schoolbox_parents.spouse_email,
schoolbox_parents.spouse_carer_number,
contact.firstname,
contact.surname



from schoolbox_parents
left join carer on schoolbox_parents.spouse_carer_number = carer.carer_number
left join contact on carer.contact_id = contact.contact_id

"

        Using conn As New System.Data.Odbc.OdbcConnection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New System.Data.Odbc.OdbcCommand(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As System.Data.Odbc.OdbcDataReader
            dr = command.ExecuteReader

            Dim i As Integer = 0
            While dr.Read()
                If Not dr.IsDBNull(0) Then
                    users.Add(New user)

                    users.Last.Delete = ""
                    users.Last.SchoolboxUserID = ""
                    users.Last.Title = ""
                    users.Last.Role = "Parents"
                    users.Last.Campus = "Senior"
                    users.Last.Password = ""
                    users.Last.AltEmail = ""
                    users.Last.Year = ""
                    users.Last.ResidentialHouse = ""
                    users.Last.EPortfolio = "N"
                    users.Last.HideContactDetails = "Y"
                    users.Last.HideTimetable = "Y"
                    users.Last.EmailAddressFromUsername = "N"
                    users.Last.UseExternalMailClient = "Y"
                    users.Last.EnableWebmailTab = "N"
                    users.Last.Superuser = "N"
                    users.Last.AccountEnabled = "Y"
                    users.Last.HomePhone = ""
                    users.Last.MobilePhone = ""
                    users.Last.WorkPhone = ""
                    users.Last.DateOfBirth = ""
                    users.Last.Address = ""
                    users.Last.Suburb = ""
                    users.Last.Postcode = ""
                    If Not dr.IsDBNull(0) Then users.Last.Username = Strings.Left(dr.GetValue(0), Strings.InStr(dr.GetValue(0), "@") - 1)
                    If Not dr.IsDBNull(1) Then users.Last.ExternalID = dr.GetValue(1)
                    If Not dr.IsDBNull(2) Then users.Last.FirstName = """" & dr.GetValue(2) & """"
                    If Not dr.IsDBNull(3) Then users.Last.Surname = """" & dr.GetValue(3) & """"


                    For Each a In studentParents
                        If users.Last.ExternalID = a.parent_id Then
                            If users.Last.ChildExternalIDs = "" Then
                                users.Last.ChildExternalIDs = a.student_id
                            Else
                                users.Last.ChildExternalIDs = users.Last.ChildExternalIDs & "," & a.student_id
                            End If
                        End If

                    Next
                    users.Last.ChildExternalIDs = """" & users.Last.ChildExternalIDs & """"

                End If


            End While

        End Using



        Dim sw As New StreamWriter(".\user.csv")
        sw.WriteLine("Delete?,Schoolbox User ID,Username,External ID,Title,First Name,Surname,Role,Campus,Password,Alt Email,Year,House,Residential House,E-Portfolio?,Hide Contact Details?,Hide Timetable?,Email Address From Username?,Use External Mail Client?,Enable Webmail Tab?,Account Enabled?,Child External IDs,Date of Birth,Home Phone,Mobile Phone,Work Phone,Address,Suburb,Postcode")
        For Each i In users
            sw.WriteLine(i.Delete & "," & i.SchoolboxUserID & "," & i.Username & "," & i.ExternalID & "," & i.Title & "," & i.FirstName & "," & i.Surname & "," & i.Role & "," & i.Campus & "," & i.Password & "," & i.AltEmail & "," & i.Year & "," & i.House & "," & i.ResidentialHouse & "," & i.EPortfolio & "," & i.HideContactDetails & "," & i.HideTimetable & "," & i.EmailAddressFromUsername & "," & i.UseExternalMailClient & "," & i.EnableWebmailTab & "," & i.AccountEnabled & "," & i.ChildExternalIDs & "," & i.DateOfBirth & "," & i.HomePhone & "," & i.MobilePhone & "," & i.WorkPhone & "," & i.Address & "," & i.Suburb & "," & i.Postcode)



        Next


    End Sub


    Function ddMMYYYY_to_yyyyMMdd(inString As String)
        ddMMYYYY_to_yyyyMMdd = Strings.Right(inString, 4) & "-" & Left(Mid(inString, Strings.InStr(inString, "/") + 1), 2) & "-" & Left(inString, InStr(inString, "/") - 1)

    End Function




    Sub timetableStructure(config As configSettings)

        Dim sep As String = ","
        Dim commandString As String
        commandString = "
SELECT        CASE WHEN substr(timetable.timetable, 6, 6) = 'Year 1' THEN 'Senior' ELSE substr(timetable.timetable, 6, 6) END AS Expr1, CONCAT(CONCAT(term.term, ' '), 
                         substr(timetable.timetable, 1, 4)) AS Expr2, term.start_date, term.end_date, term.cycle_start_day, cycle_day.day_index, period.period, period.start_time, 
                         period.end_time
FROM            OFGSODBC.TERM_GROUP, cycle_day, period_cycle_day, period, term, timetable
WHERE        (start_date > '01/01/2016') AND (end_date < '12/31/2017') AND (term_group.cycle_id = cycle_day.cycle_id) AND 
                         (cycle_day.cycle_day_id = period_cycle_day.cycle_day_id) AND (period_cycle_day.period_id = period.period_id) AND (term_group.term_id = term.term_id) AND 
                         (term.timetable_id = timetable.timetable_id)
ORDER BY timetable.timetable, term.start_date, cycle_day.day_index, period.start_time"






        Dim sw As New StreamWriter(".\timetableStructure.csv")

        Dim ConnectionString As String = config.connectionString
        Using conn As New System.Data.Odbc.OdbcConnection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New System.Data.Odbc.OdbcCommand(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As System.Data.Odbc.OdbcDataReader
            dr = command.ExecuteReader


            sw.WriteLine("Term Campus,Term Title,Term Start,Term Finish,Term Start Day Number,Period Day,Period Title,Period Start,Period Finish")

            Dim fields As Integer = dr.FieldCount - 1
            While dr.Read()
                Dim sb As New StringBuilder()

                Dim outLine As String

                outLine = (dr.GetValue(0) & "," & dr.GetValue(1) & "," & Format(dr.GetValue(2), "yyyy-MM-dd") & "," & Format(dr.GetValue(3), "yyyy-MM-dd") & "," & dr.GetValue(4) & "," & dr.GetValue(5).ToString & "," & dr.GetValue(6).ToString & "," & dr.GetValue(7).ToString & "," & dr.GetValue(8).ToString)
                sw.WriteLine(outLine)
            End While
        End Using








    End Sub


    Sub timetable(config As configSettings)
        Dim commandstring As String
        commandstring = "
SELECT DISTINCT
 substr(timetable.timetable, 6, 6) as CAMPUS1,	
CONCAT(CONCAT(term.term, ' '), substr(timetable.timetable, 1, 4)) AS Expr2,
cycle_day.day_index as DAY_NUMBER,
	period.period as PERIOD_NUMBER,
	concat(course.code,class.identifier) AS CLASS_CODE,
	class.class,
	room.code AS ROOM,
staff.staff_number

FROM period_class
INNER JOIN period_cycle_day ON period_cycle_day.period_cycle_day_id = period_class.period_cycle_day_id
INNER JOIN cycle_day ON cycle_day.cycle_day_id = period_cycle_day.cycle_day_id
INNER JOIN period ON period.period_id = period_cycle_day.period_id
INNER JOIN class ON class.class_id = period_class.class_id
INNER JOIN course ON course.course_id = class.course_id
INNER JOIN perd_cls_teacher ON perd_cls_teacher.period_class_id = period_class.period_class_id 
	AND perd_cls_teacher.is_primary = 1
INNER JOIN teacher ON teacher.teacher_id = perd_cls_teacher.teacher_id
INNER JOIN staff ON staff.contact_id = teacher.contact_id
INNER JOIN room ON room.room_id = period_class.room_id
INNER JOIN timetable ON timetable.timetable_id = period_class.timetable_id
INNER JOIN contact ON staff.contact_id = contact.contact_id
INNER JOIN term_group on cycle_day.cycle_id = term_group.cycle_id
INNER JOIN term ON term_group.term_id = term.term_id
WHERE
(
	current date BETWEEN (
	CASE
		WHEN period_class.effective_start = timetable.computed_start_date
		THEN timetable.computed_v_start_date
		ELSE period_class.effective_start
	END
	)
	AND period_class.effective_end
)
AND
(
	current date BETWEEN (
	CASE
		WHEN period_class.effective_start = timetable.computed_start_date
		THEN timetable.computed_v_start_date
		ELSE period_class.effective_start
	END
	)
	AND timetable.computed_end_date
)"

        Dim sw As New StreamWriter(".\timetable.csv")

        Dim ConnectionString As String = config.connectionString
        Using conn As New System.Data.Odbc.OdbcConnection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New System.Data.Odbc.OdbcCommand(commandstring, conn)
            command.Connection = conn
            command.CommandText = commandstring

            Dim dr As System.Data.Odbc.OdbcDataReader
            dr = command.ExecuteReader


            sw.WriteLine("Term Campus,Term Title,Period Day,Period Title,Class Code,Class Title,Class Room,Teacher Code")

            Dim fields As Integer = dr.FieldCount - 1
            While dr.Read()
                Dim sb As New StringBuilder()

                Dim outLine As String
                Dim tempStr As String

                tempStr = dr.GetValue(5)
                tempStr = Replace(tempStr, "&#039;", "'")
                tempStr = Replace(tempStr, "&amp;", "&")


                outLine = (dr.GetValue(0) & "," & dr.GetValue(1) & "," & dr.GetValue(2) & "," & dr.GetValue(3) & "," & dr.GetValue(4) & ",""" & tempStr & """," & dr.GetValue(6) & "," & dr.GetValue(7))
                sw.WriteLine(outLine)
            End While
        End Using

    End Sub

    Sub enrollment(config As configSettings)
        Dim commandstring As String
        commandstring = "
SELECT DISTINCT CONCAT(course.code, class.identifier) AS CLASS_CODE, class.class, student.student_number
FROM            OFGSODBC.CLASS_ENROLLMENT, OFGSODBC.STUDENT, class, course, academic_year
WHERE        (class_enrollment.student_id = student.student_id) AND (class_enrollment.class_id = class.class_id) AND (class.course_id = course.course_id) AND (class.academic_year_id = academic_year.academic_year_id) AND academic_year.academic_year = '" & Date.Today.Year & "'"


        Dim sw As New StreamWriter(".\enrollment.csv")

        Dim ConnectionString As String = config.connectionString
        Using conn As New System.Data.Odbc.OdbcConnection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New System.Data.Odbc.OdbcCommand(commandstring, conn)
            command.Connection = conn
            command.CommandText = commandstring

            Dim dr As System.Data.Odbc.OdbcDataReader
            dr = command.ExecuteReader


            sw.WriteLine("Class Code,Class Title,Student Code")

            Dim fields As Integer = dr.FieldCount - 1
            While dr.Read()
                Dim sb As New StringBuilder()

                Dim outLine As String
                Dim tempStr As String

                tempStr = dr.GetValue(1)
                tempStr = Replace(tempStr, "&#039;", "'")
                tempStr = Replace(tempStr, "&amp;", "&")

                outLine = (dr.GetValue(0) & ",""" & tempStr & """," & dr.GetValue(2))
                sw.WriteLine(outLine)
            End While
        End Using




        '&#039;
        '&#039;
        '&amp;



    End Sub


    Sub upload(host As String, userName As String, pass As String, rsa As String)
        Try
            ' Setup session options
            Dim sessionOptions As New SessionOptions
            With sessionOptions
                .Protocol = Protocol.Sftp
                .HostName = host
                .UserName = userName
                .Password = pass
                .SshHostKeyFingerprint = rsa '"ssh-rsa 2048 xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx"
            End With

            Using session As New Session
                ' Connect
                session.Open(sessionOptions)

                ' Upload files
                Dim transferOptions As New TransferOptions
                transferOptions.TransferMode = TransferMode.Binary

                Dim transferResult As TransferOperationResult
                transferResult = session.PutFiles(".\*.csv", "./", False, transferOptions)

                ' Throw on any error
                transferResult.Check()

                ' Print results
                For Each transfer In transferResult.Transfers
                    Console.WriteLine("Upload of {0} succeeded", transfer.FileName)
                Next
            End Using


        Catch e As Exception
            Console.WriteLine("Error: {0}", e)

        End Try

    End Sub

    Sub uploadFiles(config As configSettings)
        For Each i In config.uploadServers
            upload(i.host, i.userName, i.pass, i.rsa)
        Next
    End Sub

    Private Function readConfig()
        Dim config As New configSettings()
        config.uploadServers = New List(Of uploadServer)

        Try
            ' Open the file using a stream reader.
            Dim directory As String = My.Application.Info.DirectoryPath



            Using sr As New StreamReader(directory & "\config.ini")
                Dim line As String
                While Not sr.EndOfStream
                    line = sr.ReadLine

                    Select Case True
                        Case Left(line, 17) = "connectionstring="
                            config.connectionString = Mid(line, 18)
                        Case Left(line, 13) = "uploadserver="
                            line = Mid(line, 14)
                            Dim split As String() = line.Split(";")
                            config.uploadServers.Add(New uploadServer)
                            config.uploadServers.Last.host = split(0)
                            config.uploadServers.Last.userName = split(1)
                            config.uploadServers.Last.pass = split(2)
                            config.uploadServers.Last.rsa = split(3)
                        Case Left(line, 19) = "studentemaildomain="
                            config.studentEmailDomain = Mid(line, 20)
                    End Select

                End While

                readConfig = config
            End Using

        Catch e As Exception
            MsgBox(e.Message)
        End Try
    End Function



End Module


