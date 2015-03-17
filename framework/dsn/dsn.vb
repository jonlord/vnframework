Public Class dsn
    Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv As Integer, ByVal fDirection As Short, ByVal szDSN As String, ByVal cbDSNMax As Short, ByRef pcbDSN As Short, ByVal szDescription As String, ByVal cbDescriptionMax As Short, ByRef pcbDescription As Short) As Short
    Private Declare Function SQLAllocEnv Lib "ODBC32.DLL" (ByRef env As Integer) As Short

    Const SQL_SUCCESS As Integer = 0
    Const SQL_FETCH_NEXT As Integer = 1
    Private Const ODBC_ADD_DSN As Short = 1 ' Add user data source
    Private Const ODBC_CONFIG_DSN As Short = 2 ' Configure (edit) data source
    Private Const ODBC_REMOVE_DSN As Short = 3 ' Remove data source
    Private Const ODBC_ADD_SYS_DSN As Short = 4 'Add system data source
    Private Const vbAPINull As Integer = 0 ' NULL Pointer

    Public Shared Sub FetchDSNs()
        Dim ReturnValue As Short
        Dim DSNName As String
        Dim DriverName As String
        Dim DSNNameLen As Short
        Dim DriverNameLen As Short
        Dim SQLEnv As Integer 'handle to the environment

        If SQLAllocEnv(SQLEnv) <> -1 Then
            Do Until ReturnValue <> SQL_SUCCESS
                DSNName = Space(1024)
                DriverName = Space(1024)
                ReturnValue = SQLDataSources(SQLEnv, SQL_FETCH_NEXT, DSNName, 1024, DSNNameLen, DriverName, 1024, DriverNameLen)
                DSNName = Left(DSNName, DSNNameLen)
                DriverName = Left(DriverName, DriverNameLen)

                If DSNName <> Space(DSNNameLen) And DSNName = "FA" Then
                    Debug.WriteLine(DSNName)
                    Debug.WriteLine(DriverName)
                End If
            Loop
        End If
    End Sub
End Class