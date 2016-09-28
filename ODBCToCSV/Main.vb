Imports System
Imports System.IO


Imports System.Data
Imports System.Data.OleDb
Imports System.Text
Imports System.Object
Imports System.Web
Imports System.Text.RegularExpressions


Module Main

    Public Sub Main(ByVal sArgs() As String)
        Const seperator = ","
        If sArgs.Length = 0 Then                'If there are no arguments
            Console.WriteLine("Type query string on stdin") '

        Else
            Dim i As Integer = 0
            While i < sArgs.Length             'So with each argument
                WriteToTextFile(seperator, sArgs(i))
                i = i + 1                       'Increment to the next argument
            End While
        End If



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




    Public Sub WriteToTextFile(ByVal separator As String, commandStringIn As String)

        Dim ConnStr As String = readConnectionString()
        Dim commandString As String = commandStringIn
        Dim sep As String = separator
        Dim cleanString As String = Regex.Replace(commandString, "[^A-Za-z0-9\-/]", "")


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




End Module
