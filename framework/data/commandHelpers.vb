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
End Module