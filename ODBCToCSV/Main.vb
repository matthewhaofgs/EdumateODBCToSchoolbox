﻿Imports System
Imports System.IO


Imports System.Data
Imports System.Data.OleDb
Imports System.Text
Imports System.Object
Imports System.Web
Imports System.Text.RegularExpressions

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


Module Main

    Public Sub Main(ByVal sArgs() As String)



        '      Call user()
        Call timetableStructure()

    End Sub


    Private Function readConnectionString()
        Try
            ' Open the file using a stream reader.
            Dim directory As String = My.Application.Info.DirectoryPath
            Using sr As New StreamReader(directory & "\connectionString.txt")
                Dim line As String
                line = sr.ReadToEnd()
                readConnectionString = line
            End Using
        Catch e As Exception
            MsgBox(e.Message)
        End Try
    End Function




    Public Sub WriteToTextFile(ByVal separator As String, commandStringIn As String, outFileName As String)

        Dim ConnStr As String = readConnectionString()
        Dim commandString As String = commandStringIn
        Dim cleanString As String
        Dim sep As String = separator

        If outFileName = "" Then
            cleanString = Regex.Replace(commandString, "[^A-Za-z0-9\-/]", "")
        Else
            cleanString = outFileName
        End If


        Dim sw As New StreamWriter(".\" & cleanString & ".csv")

        Dim ConnectionString As String = readConnectionString()
        Using conn As New System.Data.Odbc.OdbcConnection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New System.Data.Odbc.OdbcCommand(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As System.Data.Odbc.OdbcDataReader
            dr = command.ExecuteReader




            Dim fields As Integer = dr.FieldCount - 1
            While dr.Read()
                Dim sb As New StringBuilder()
                Dim i As Integer = 0
                While i <= fields
                    If i <> fields Then
                        sep = sep
                    Else
                        ' sep = ""
                    End If
                    sb.Append(dr(i).ToString() + sep)
                    i += 1
                End While
                sw.WriteLine(sb.ToString())
            End While
        End Using


    End Sub

    Sub user()

        ' Students ****************
        Dim ConnStr As String = readConnectionString()
        Dim commandString As String = "
select
username,
student_number,
firstname,
surname,
birthdate
from schoolbox_students
"
        Dim users As New List(Of user)


        Dim ConnectionString As String = readConnectionString()
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
                users.Last.Role = "Senior Students"
                users.Last.Campus = "Senior"
                users.Last.Password = ""
                users.Last.AltEmail = ""
                users.Last.Year = ""
                users.Last.House = ""
                users.Last.ResidentialHouse = ""
                users.Last.EPortfolio = "Y"
                users.Last.HideContactDetails = "Y"
                users.Last.HideTimetable = "N"
                users.Last.EmailAddressFromUsername = "Y"
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
                If Not dr.IsDBNull(2) Then users.Last.FirstName = dr.GetValue(2)
                If Not dr.IsDBNull(3) Then users.Last.Surname = dr.GetValue(3)
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
                If Not dr.IsDBNull(3) Then users.Last.FirstName = dr.GetValue(3)
                If Not dr.IsDBNull(4) Then users.Last.Surname = dr.GetValue(4)
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
                users.Last.EmailAddressFromUsername = "Y"
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
                If Not dr.IsDBNull(2) Then users.Last.FirstName = dr.GetValue(2)
                If Not dr.IsDBNull(3) Then users.Last.Surname = dr.GetValue(3)


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


        Dim sw As New StreamWriter(".\user.csv")
        sw.WriteLine("Delete?,Schoolbox User ID,Username,External ID,Title,First Name,Surname,Role,Campus,Password,Alt Email,Year,House,Residential House,E-Portfolio?,Hide Contact Details?,Hide Timetable?,Email Address From Username?,Use External Mail Client?,Enable Webmail Tab?,Superuser?,Account Enabled?,Child External IDs,Date of Birth,Home Phone,Mobile Phone,Work Phone,Address,Suburb,Postcode")
        For Each i In users
            sw.WriteLine(i.Delete & "," & i.SchoolboxUserID & "," & i.Username & "," & i.ExternalID & "," & i.Title & "," & i.FirstName & "," & i.Surname & "," & i.Role & "," & i.Campus & "," & i.Password & "," & i.AltEmail & "," & i.Year & "," & i.House & "," & i.ResidentialHouse & "," & i.EPortfolio & "," & i.HideContactDetails & "," & i.HideTimetable & "," & i.EmailAddressFromUsername & "," & i.UseExternalMailClient & "," & i.EnableWebmailTab & "," & i.Superuser & "," & i.AccountEnabled & "," & i.ChildExternalIDs & "," & i.DateOfBirth & "," & i.HomePhone & "," & i.MobilePhone & "," & i.WorkPhone & "," & i.Address & "," & i.Suburb & "," & i.Postcode)



        Next


    End Sub


    Function ddMMYYYY_to_yyyyMMdd(inString As String)
        ddMMYYYY_to_yyyyMMdd = Strings.Right(inString, 4) & "-" & Left(Mid(inString, Strings.InStr(inString, "/") + 1), 2) & "-" & Left(inString, InStr(inString, "/") - 1)

    End Function




    Sub timetableStructure()

        Dim sep As String = ","
        Dim commandString As String
        commandString = "
SELECT        substr(timetable.timetable, 6, 6) AS Expr1, CONCAT(CONCAT(term.term, ' '), substr(timetable.timetable, 1, 4)) AS Expr2, term.start_date, term.end_date, 
                         term.cycle_start_day, cycle_day.day_index, period.period, period.start_time, period.end_time
FROM            OFGSODBC.TERM_GROUP, cycle_day, period_cycle_day, period, term, timetable
WHERE        (start_date > '01/01/2016') AND (end_date < '12/31/2017') AND (term_group.cycle_id = cycle_day.cycle_id) AND 
                         (cycle_day.cycle_day_id = period_cycle_day.cycle_day_id) AND (period_cycle_day.period_id = period.period_id) AND (term_group.term_id = term.term_id) AND 
                         (term.timetable_id = timetable.timetable_id)
ORDER BY timetable.timetable, term.start_date, cycle_day.day_index, period.start_time"






        Dim sw As New StreamWriter(".\timetableStructure.csv")

        Dim ConnectionString As String = readConnectionString()
        Using conn As New System.Data.Odbc.OdbcConnection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New System.Data.Odbc.OdbcCommand(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As System.Data.Odbc.OdbcDataReader
            dr = command.ExecuteReader




            Dim fields As Integer = dr.FieldCount - 1
            While dr.Read()
                Dim sb As New StringBuilder()

                Dim outLine As String

                outLine = (dr.GetValue(0) & "," & dr.GetValue(1) & "," & Format(dr.GetValue(2), "yyyy-MM-dd") & "," & Format(dr.GetValue(3), "yyyy-MM-dd") & "," & dr.GetValue(4) & "," & dr.GetValue(5).ToString & "," & dr.GetValue(6).ToString & "," & dr.GetValue(7).ToString & "," & dr.GetValue(8).ToString)
                sw.WriteLine(outLine)
            End While
        End Using








    End Sub


End Module
