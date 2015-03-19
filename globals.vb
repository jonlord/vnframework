Imports System.Data.SqlClient
Imports System.Data.Odbc
''' <summary>
''' Some global functions to reduce passing around variables
''' </summary>
Public Module globals
    Private _connection As OdbcConnection
    Private _sqlconnection As SqlConnection
    ''' <summary>
    ''' Get a Global ODBC Connection
    ''' </summary>
    ''' <returns>ODBC Connection</returns>
    Public Function getGlobalConnection() As OdbcConnection
        Return _connection
    End Function
    ''' <summary>
    ''' Sets a global ODBC Connection
    ''' </summary>
    ''' <param name="connection">ODBC Connection to set</param>
    Public Sub setGlobalConnection(connection As OdbcConnection)
        _connection = connection
    End Sub
    Public Function getGlobalSQLConnection() As SqlConnection
        Return _sqlconnection
    End Function
    ''' <summary>
    ''' Sets a global ODBC Connection
    ''' </summary>
    ''' <param name="connection">ODBC Connection to set</param>
    Public Sub setGlobalSQLConnection(connection As SqlConnection)
        _sqlconnection = connection
    End Sub
End Module