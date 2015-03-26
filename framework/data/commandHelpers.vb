Imports System.Data.Odbc
Imports System.Data.OleDb
Imports MySql.Data.MySqlClient

''' <summary>
''' Library of functions to make data retreival from relational databases easier
''' </summary>
Public Module commandHelpers
    ''' <summary>
    ''' Retrieve a scalar value from a database
    ''' </summary>
    ''' <param name="command">Command to use</param>
    ''' <returns>Scalar value</returns>
    Function ExecuteScalar(ByVal command As OleDbCommand) As Object
        command.CommandText = sql.parseToSQLServer(command.CommandText) 'Let's make the query compatible with SQL Server
        logs.writeToLog(command.CommandText) 'This line is optional, I like having my queries "logged" someplace I can check for errors, for development purposes I use Debug.Print
        Return command.ExecuteScalar
    End Function
    ''' <summary>
    ''' Retreive a set of records from a database
    ''' </summary>
    ''' <param name="command">Command to use</param>
    ''' <returns>Datareader containing the retreived records</returns>
    Function ExecuteReader(ByVal command As OleDbCommand) As OleDbDataReader
        command.CommandText = sql.parseToSQLServer(command.CommandText) 'Let's make the query compatible with SQL Server
        logs.writeToLog(command.CommandText) 'This line is optional, I like having my queries "logged" someplace I can check for errors, for development purposes I use Debug.Print
        Return command.ExecuteReader
    End Function
    ''' <summary>
    ''' Execute a query against a database
    ''' </summary>
    ''' <param name="command">Command to use</param>
    ''' <returns>Records affected</returns>
    Function ExecuteNonQuery(ByVal command As OleDbCommand) As Integer
        command.CommandText = sql.parseToSQLServer(command.CommandText) 'Let's make the query compatible with SQL Server
        logs.writeToLog(command.CommandText) 'This line is optional, I like having my queries "logged" someplace I can check for errors, for development purposes I use Debug.Print
        Return command.ExecuteNonQuery
    End Function

    ''' <summary>
    ''' Execute a query against a database
    ''' </summary>
    ''' <param name="command">Command to use</param>
    ''' <returns>Records affected</returns>
    Function ExecuteNonQuery(ByVal command As MySqlCommand) As Integer
        command.CommandText = sql.parseToSQLServer(command.CommandText) 'Let's make the query compatible with SQL Server
        logs.writeToLog(command.CommandText) 'This line is optional, I like having my queries "logged" someplace I can check for errors, for development purposes I use Debug.Print
        Return command.ExecuteNonQuery()
    End Function
    ''' <summary>
    ''' Retrieve a scalar value from a database
    ''' </summary>
    ''' <param name="command">Command to use</param>
    ''' <returns>Scalar value</returns>
    Function ExecuteScalar(ByVal command As MySqlCommand) As Object
        command.CommandText = sql.parseToSQLServer(command.CommandText) 'Let's make the query compatible with SQL Server
        logs.writeToLog(command.CommandText) 'This line is optional, I like having my queries "logged" someplace I can check for errors, for development purposes I use Debug.Print
        Return command.ExecuteScalar
    End Function
    ''' <summary>
    ''' Retreive a set of records from a database
    ''' </summary>
    ''' <param name="command">Command to use</param>
    ''' <returns>Datareader containing the retreived records</returns>
    Function ExecuteReader(ByVal command As MySqlCommand) As MySqlDataReader
        command.CommandText = sql.parseToSQLServer(command.CommandText) 'Let's make the query compatible with SQL Server
        logs.writeToLog(command.CommandText) 'This line is optional, I like having my queries "logged" someplace I can check for errors, for development purposes I use Debug.Print
        Return command.ExecuteReader
    End Function

    ''' <summary>
    ''' Retrieve a scalar value from a database
    ''' </summary>
    ''' <param name="command">Command to use</param>
    ''' <returns>Scalar value</returns>
    Function ExecuteScalar(ByVal command As OdbcCommand) As Object
        command.CommandText = sql.parseToSQLServer(command.CommandText) 'Let's make the query compatible with SQL Server
        logs.writeToLog(command.CommandText) 'This line is optional, I like having my queries "logged" someplace I can check for errors, for development purposes I use Debug.Print
        Return command.ExecuteScalar
    End Function
    ''' <summary>
    ''' Retrieve a scalar value from a database
    ''' </summary>
    ''' <param name="commandText">Query to run against database</param>
    ''' <param name="connection">Connection to use</param>
    ''' <returns>Scalar value</returns>
    Function ExecuteScalar(ByVal commandText As String, connection As OdbcConnection) As Object
        commandText = sql.parseToSQLServer(commandText) 'Let's make the query compatible with SQL Server
        logs.writeToLog(commandText) 'This line is optional, I like having my queries "logged" someplace I can check for errors, for development purposes I use Debug.Print
        If connection Is Nothing Then connection = getGlobalConnection() 'If we did not set a connection, let's get the default one
        If connection.State <> ConnectionState.Open Then connection.Open()
        Using cmd As New OdbcCommand(commandText, connection)
            Return cmd.ExecuteScalar()
        End Using
    End Function
    ''' <summary>
    ''' Execute a query against a database
    ''' </summary>
    ''' <param name="command">Command to use</param>
    ''' <returns>Records affected</returns>
    Function ExecuteNonQuery(ByVal command As OdbcCommand) As Integer
        command.CommandText = sql.parseToSQLServer(command.CommandText) 'Let's make the query compatible with SQL Server
        logs.writeToLog(command.CommandText) 'This line is optional, I like having my queries "logged" someplace I can check for errors, for development purposes I use Debug.Print
        Return command.ExecuteNonQuery()
    End Function
    ''' <summary>
    ''' Execute a query against a database
    ''' </summary>
    ''' <param name="commandText">Query to run against database</param>
    ''' <param name="connection">Connection to use</param>
    ''' <returns>Records affected</returns>
    Function ExecuteNonQuery(ByVal commandText As String, connection As OdbcConnection) As Integer
        commandText = sql.parseToSQLServer(commandText) 'Let's make the query compatible with SQL Server
        logs.writeToLog(commandText) 'This line is optional, I like having my queries "logged" someplace I can check for errors, for development purposes I use Debug.Print
        If connection Is Nothing Then connection = getGlobalConnection() 'If we did not set a connection, let's get the default one
        If connection.State <> ConnectionState.Open Then connection.Open()
        Using cmd As New OdbcCommand(commandText, connection)
            Return cmd.ExecuteNonQuery()
        End Using
    End Function
    ''' <summary>
    ''' Retreive a set of records from a database
    ''' </summary>
    ''' <param name="command">Command to use</param>
    ''' <returns>Datareader containing the retreived records</returns>
    Function ExecuteReader(ByVal command As OdbcCommand) As OdbcDataReader
        command.CommandText = sql.parseToSQLServer(command.CommandText) 'Let's make the query compatible with SQL Server
        logs.writeToLog(command.CommandText) 'This line is optional, I like having my queries "logged" someplace I can check for errors, for development purposes I use Debug.Print
        Return command.ExecuteReader
    End Function
    ''' <summary>
    ''' Retreive a set of records from a database
    ''' </summary>
    ''' <param name="commandText">Query to run against database</param>
    ''' <param name="connection">Connection to use</param>
    ''' <returns>Datareader containing the retreived records</returns>
    Function ExecuteReader(ByVal commandText As String, connection As OdbcConnection) As OdbcDataReader
        commandText = sql.parseToSQLServer(commandText) 'Let's make the query compatible with SQL Server
        logs.writeToLog(commandText) 'This line is optional, I like having my queries "logged" someplace I can check for errors, for development purposes I use Debug.Print
        If connection Is Nothing Then connection = getGlobalConnection() 'If we did not set a connection, let's get the default one
        If connection.State <> ConnectionState.Open Then connection.Open()
        Using cmd As New OdbcCommand(commandText, connection) With {.CommandTimeout = 90}
            Return cmd.ExecuteReader()
        End Using
    End Function

    ''' <summary>
    ''' Retrieve a scalar value from a database
    ''' </summary>
    ''' <param name="commandText">Query to run against database</param>
    ''' <returns>Scalar Values</returns>
    Function ExecuteScalar(ByVal commandText As String) As Object
        Return ExecuteScalar(commandText, getGlobalConnection())
    End Function
    ''' <summary>
    ''' Execute a query against a database
    ''' </summary>
    ''' <param name="commandText">Query to run against database</param>
    ''' <returns>Records affected</returns>
    Function ExecuteNonQuery(ByVal commandText As String) As Integer
        Return ExecuteNonQuery(commandText, getGlobalConnection())
    End Function
    ''' <summary>
    ''' Retreive a set of records from a database
    ''' </summary>
    ''' <param name="commandText">Query to run against database</param>
    ''' <returns>Datareader containing the retreived records</returns>
    Function ExecuteReader(ByVal commandText As String) As OdbcDataReader
        Return ExecuteReader(commandText, getGlobalConnection())
    End Function

    ''' <summary>
    '''
    ''' </summary>
    ''' <param name="queries"></param>
    ''' <param name="cancelOnError"></param>
    ''' <param name="useTran"></param>
    ''' <param name="cnx"></param>
    ''' <returns></returns>
    Public Function procedure(ByVal queries As List(Of String), Optional ByVal cancelOnError As Boolean = True, Optional ByVal useTran As Boolean = True, Optional cnx As OdbcConnection = Nothing) As Integer
        Return procedure(queries.ToArray, cancelOnError, useTran, cnx)
    End Function
    ''' <summary>
    '''
    ''' </summary>
    ''' <param name="queries"></param>
    ''' <param name="cancelOnError"></param>
    ''' <param name="useTran"></param>
    ''' <param name="cnx"></param>
    ''' <returns></returns>
    Public Function procedure(ByVal queries As ArrayList, Optional ByVal cancelOnError As Boolean = True, Optional ByVal useTran As Boolean = True, Optional cnx As OdbcConnection = Nothing) As Integer
        Return procedure(queries.ToArray, cancelOnError, useTran, cnx)
    End Function
    ''' <summary>
    ''' Executes a SET of queries one by one
    ''' </summary>
    ''' <param name="queries">Set of Queries</param>
    ''' <param name="cancelOnError">Cancel the process if an error occurs</param>
    ''' <param name="useTran">Use a transaction</param>
    ''' <param name="connection">Connection to use</param>
    ''' <returns>1 if successful or an error code if it failts</returns>
    Public Function procedure(ByVal queries As String(), Optional ByVal cancelOnError As Boolean = True, Optional ByVal useTran As Boolean = True, Optional connection As OdbcConnection = Nothing) As Integer
        Dim i As Integer
        Dim tran As OdbcTransaction = Nothing
        Try
            Dim rSQL(queries.Count - 1) As String
            Dim cmd As New OdbcCommand("")

            If connection Is Nothing Then
                cmd.Connection = getGlobalConnection()
                If useTran Then tran = cmd.Connection.BeginTransaction
            Else
                cmd.Connection = connection
                If useTran Then tran = connection.BeginTransaction
            End If

            If useTran Then cmd.Transaction = tran
            cmd.CommandTimeout = 600
            For i = 0 To queries.Count - 1
                If queries(i) = "" Then Continue For
                For j As Integer = 0 To i
                    queries(i) = Replace(queries(i), String.Format("@@rSQL{0}@@", j), rSQL(j)).Trim()
                Next
                cmd.CommandText = queries(i)
                If queries(i).ToUpper.StartsWith("SELECT") Then
                    Using dr As OdbcDataReader = ExecuteReader(cmd)
                        If dr.Read Then
                            If dr.GetName(0).ToUpper = "VERIFY" Then 'Columns called VERIFY are used to validate if a query was successful
                                If CLng(dr(0)) = 0 Then
                                    If useTran Then tran.Rollback()
                                    Return 0
                                End If
                            End If
                            rSQL(i) = dr(0)
                        End If
                        dr.Close()
                    End Using
                Else
                    ExecuteNonQuery(cmd)
                End If
            Next
            If useTran Then tran.Commit()
            Return 1
        Catch e1 As OdbcException
            If Not cancelOnError Then
                For j As Integer = 1 To queries.Length - 1
                    queries(j - 1) = queries(j)
                Next
                ReDim Preserve queries(queries.Length - 2)
                Return procedure(queries, cancelOnError)
            End If
            Try
                If Not IsNothing(tran) And useTran Then tran.Rollback()
            Catch e2 As Exception
                logs.writeToLog(e2.Message)
                logs.writeToLog(e2.StackTrace)
            End Try
            Dim errMsg As String = vnframework.sqlserver.cleanError(e1.Message)

            logs.writeToErrorLog(errMsg)
            logs.writeToLog(errMsg)
            logs.writeToLog(errMsg)
            Return e1.ErrorCode
        End Try
        Return 0
    End Function

    ''' <summary>
    ''' Executes a SET of queries one by one
    ''' </summary>
    ''' <param name="queries">Set of Queries</param>
    ''' <param name="mysqlconnection">Connection to use</param>
    ''' <param name="cancelOnError">Cancel the process if an error occurs</param>
    ''' <param name="useTran">Use a transaction</param>
    ''' <returns>1 if successful or an error code if it failts</returns>
    Function procedureMySQL(ByVal queries As String(), MySQLConnection As MySqlConnection, Optional ByVal cancelOnError As Boolean = True, Optional ByVal useTran As Boolean = True) As Integer
        Dim i As Integer
        Dim tran As MySqlTransaction = Nothing
        Try
            If MySQLConnection.State <> ConnectionState.Open Then MySQLConnection.Open()
            Dim rSQL(queries.Count - 1) As String
            If useTran Then tran = MySQLConnection.BeginTransaction
            Using cmd As New MySqlCommand("", MySQLConnection)
                If useTran Then cmd.Transaction = tran
                cmd.CommandTimeout = 600
                For i = 0 To queries.Count - 1
                    If queries(i) = "" Then Continue For
                    For j As Integer = 0 To i
                        queries(i) = Replace(queries(i), String.Format("@@rSQL{0}@@", j), rSQL(j))
                    Next
                    cmd.CommandText = queries(i)
                    If queries(i).ToUpper.StartsWith("SELECT") Then
                        Using dr As MySqlDataReader = ExecuteReader(cmd)
                            If dr.Read Then
                                If dr.GetName(0).ToUpper = "VERIFY" And CInt(dr(0)) = 0 Then
                                    If useTran Then tran.Rollback()
                                    Return 0
                                End If
                                rSQL(i) = dr(0)
                            End If
                            dr.Close()
                        End Using
                    Else
                        ExecuteNonQuery(cmd)
                    End If
                Next
            End Using
            If useTran Then tran.Commit()
            MySQLConnection.Close()
            Return 1
        Catch e1 As MySqlException
            If Not cancelOnError Then
                For j As Integer = 1 To queries.Length - 1
                    queries(j - 1) = queries(j)
                Next
                ReDim Preserve queries(queries.Length - 2)
                Return procedureMySQL(queries, MySQLConnection, cancelOnError)
            End If
            Try
                If Not IsNothing(tran) And useTran Then tran.Rollback()
            Catch e2 As Exception
                logs.writeToLog(e2.Message)
                logs.writeToLog(e2.StackTrace)
            End Try
            logs.writeToErrorLog(e1.Message)
            logs.writeToLog(e1.Message)
            logs.writeToLog(e1.StackTrace)
            Return e1.ErrorCode
        End Try
        Return 0
    End Function

    ''' <summary>
    ''' Executes a SET of queries one by one
    ''' </summary>
    ''' <param name="queries">Set of Queries</param>
    ''' <param name="mysqlconnection">Connection to use</param>
    ''' <param name="cancelOnError">Cancel the process if an error occurs</param>
    ''' <param name="useTran">Use a transaction</param>
    ''' <returns>1 if successful or an error code if it failts</returns>
    Public Function procedureMySQL(ByVal queries As ArrayList, MySQLConnection As MySqlConnection, Optional ByVal cancelOnError As Boolean = True, Optional ByVal useTran As Boolean = True) As Integer
        Return procedureMySQL(queries.ToArray, MySQLConnection, cancelOnError, useTran)
    End Function
End Module