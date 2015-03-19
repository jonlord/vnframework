Imports System.IO
Imports System.Drawing
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

''' <summary>
''' SQL Server compatibility functions
''' </summary>
Public Class sqlserver
    ''' <summary>
    ''' Gets de function for the current date
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function nowFunction() As String
        nowFunction = " GETDATE() "
        'If Format() = "YYYY-MM-DD" Then nowFunction = " CONVERT(VARCHAR(10), GETDATE(), 120) "
    End Function
    ''' <summary>
    ''' Builds a SQL Snippet that checks if a value is null and returns another one if it is
    ''' </summary>
    ''' <param name="field">Field to be evaluated</param>
    ''' <param name="valueIfNull">Value to return if its null</param>
    ''' <param name="valueIfNotNull">Value to return if its not null</param>
    ''' <returns>SQL Snippet applying a null validation</returns>
    Public Shared Function ifNull(ByVal field As String, Optional ByVal valueIfNull As String = "0", Optional ByVal valueIfNotNull As String = "") As String
        Return String.Format(" ISNULL({0},{1}) ", field, valueIfNull)
    End Function
    ''' <summary>
    ''' Builds a SQL Snippet for an inline IF validation
    ''' </summary>
    ''' <param name="condition">Condition to evaluate</param>
    ''' <param name="trueValue">Value to return if is is True</param>
    ''' <param name="falseValue">Value to return if is is False</param>
    ''' <returns>SQL Snippet representing trueValue when condition is true, falseValue when condition is false</returns>
    Public Shared Function iifFunction(ByVal condition As String, ByVal trueValue As String, ByVal falseValue As String) As String
        Return String.Format(" CASE WHEN {0} THEN {1} ELSE {2} END ", condition, trueValue, falseValue)
    End Function
    ''' <summary>
    ''' Builds SQL Snippet for the Minute Representation for SQL Server
    ''' </summary>
    ''' <returns>SQL Snippet for the Minute Representation for SQL Server</returns>
    Public Shared Function dateAddMinute() As String
        Return "Minute"
    End Function
    ''' <summary>
    ''' Build SQL Snippet for the Day Representation for SQL Server
    ''' </summary>
    ''' <returns>SQL Snippet for the Day Representation for SQL Server</returns>
    Public Shared Function dateAddDay() As String
        Return "Day"
    End Function
    ''' <summary>
    ''' Builds SQL Snippet for the Month Representation for SQL Server
    ''' </summary>
    ''' <returns>SQL Snippet for the Month Representation for SQL Server</returns>
    Public Shared Function dateAddMonth() As String
        Return "Month"
    End Function

    'Database Errors used in the cleanError Function
    Private Const ERRORNULLFIELD As String = "\bCannot insert the value NULL into column \'(\w+)\', table \'(\w+.\w+.\w+)\'; column does not allow nulls. INSERT fails."
    Private Const ERRORUNCLOSEDQUOTATION As String = "\bUnclosed quotation mark after the character string"
    Private Const ERRORARITHMETICOVERFLOW As String = "Arithmetic overflow error converting expression to data type int"
    Private Const ERRORSTRINGTRUNCATED As String = "String Or binary data would be truncated."
    Private Const ERRORFOREIGNKEY As String = "The ALTER TABLE statement conflicted With the FOREIGN KEY constraint"
    Private Const ERRORMISSINGFIELD As String = "Invalid column name"
    Private Const ERRORMISSINGTABLE As String = "Invalid Object name"
    Private Const ERRORFIELDCOUNT0 As String = "FIELDCOUNT0"
    Private Const ERRORFOREIGNKEYINSERT As String = "\bThe INSERT statement conflicted With the FOREIGN KEY constraint \""(\w+)\"". The conflict occurred In databse \""(\w+)\"", table \""(\w+).(\w+)\"", column \'(\w+)\'."
    Public Const ERRORPRIMARYKEY As String = "Violation of PRIMARY KEY constraint"

    ''' <summary>
    ''' Receives a Database produced error message and makes it user friendly
    ''' </summary>
    ''' <param name="errorMessage">Error Message to make user friendly</param>
    ''' <returns>User friendly version of the error message</returns>
    Shared Function cleanError(errorMessage As String) As String
        'Lets remove the SQL Server standard error message sections
        cleanError = errorMessage.Replace("[Microsoft]", "").Replace("[SQL Server]", "").Replace("[SQL Server Native Client]", "").Replace("[SQL Server Native Client 10.0]", "").Replace("[SQL Server Native Client 11.0]", "").Replace("[SQL Native Client]", "").Replace("Error [42S02]", "").Replace("Error [22003]", "").Replace("Error [42000]", "").Replace("Error [HYT00]", "").Replace("Error [23000]", "").Replace("Error [01000]", "").Replace("Error [22001]", "").Replace("Error [42S22]", "").Replace("  ", "").Replace("The transaction ended In the trigger.", "").Replace("The batch has been aborted", "").Trim().Replace("The statement has been terminated.", "").Trim()

        'Lets replace common errors for more user friendly errors
        cleanError = Regex.Replace(cleanError, ERRORNULLFIELD, REPLACENULLFIELD)
        cleanError = Regex.Replace(cleanError, ERRORUNCLOSEDQUOTATION, REPLACEUNCLOSEDQUOTATION)
        If cleanError.ToLower.IndexOf(ERRORARITHMETICOVERFLOW.ToLower) > -1 Then _
            cleanError = cleanError.Replace(ERRORARITHMETICOVERFLOW, REPLACEARITHMETICOVERFLOW)
        If cleanError.ToLower.IndexOf(ERRORSTRINGTRUNCATED.ToLower) > -1 Then _
            cleanError = cleanError.Replace(ERRORSTRINGTRUNCATED, REPLACESTRINGTRUNCATED)
        If cleanError.ToLower.IndexOf(ERRORFOREIGNKEY.ToLower) > -1 Then _
            cleanError = cleanError.Replace(ERRORFOREIGNKEY, "").Replace("'", "").Replace(",", "").Replace("""", "").Replace("table", "").Replace(".", "").Replace("dbo", "").Replace("table", "").Replace("column", "").Replace("The conflict occurred in database", vbNewLine & vbNewLine) & REPLACEFOREIGNKEY
        If cleanError.ToLower.IndexOf(ERRORMISSINGFIELD.ToLower) > -1 Then _
            cleanError = cleanError.Replace(ERRORMISSINGFIELD, REPLACEMISSINGFIELD)
        If cleanError.ToLower.IndexOf(ERRORMISSINGTABLE.ToLower) > -1 Then _
            cleanError = cleanError.Replace(ERRORMISSINGTABLE, REPLACEMISSINGTABLE)
        If cleanError.ToLower.IndexOf(ERRORFIELDCOUNT0.ToLower) > -1 Then _
            cleanError = cleanError.Replace(ERRORFIELDCOUNT0, REPLACEFIELDCOUNT0)
        cleanError = Regex.Replace(cleanError, ERRORFOREIGNKEYINSERT, REPLACEFOREIGNKEYINSERT)

        'Replace double spaces
        cleanError = cleanError.Replace("  ", " ").Trim()

        'Do some custom replacements
        cleanError = cleanErrorCustom(cleanError)
    End Function

    ''' <summary>
    ''' Helper to cleanError; allows for custom messages; useful in adding constraint errors
    ''' </summary>
    ''' <param name="errorMessage">Error Message to make user friendly</param>
    ''' <returns>User friendly version of the error message</returns>
    Private Shared Function cleanErrorCustom(errorMessage As String) As String
        If errorMessage.IndexOf("chkdcto") > -1 And errorMessage.IndexOf("statement conflicted") > -1 Then _
            errorMessage = "Ha introducido un precio menor al costo asociado"
        cleanErrorCustom = errorMessage
    End Function

    ''' <summary>
    ''' Gets the column schema of a table by name
    ''' </summary>
    ''' <param name="cnx">ODBC Connection to usr</param>
    ''' <param name="tableName">Name of table to get column schema for</param>
    ''' <returns></returns>
    Shared Function getSchema(cnx As OdbcConnection, tableName As String) As Dictionary(Of String, String)
        Dim query As String = String.Format("SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='{0}' ", tableName.escapeCharacters)
        Dim cmdSource As New OdbcCommand(query, cnx)
        Dim dataReader As OdbcDataReader = cmdSource.ExecuteReader()

        Dim fields As New Dictionary(Of String, String)
        Do While dataReader.Read()
            fields.Add(dataReader(0), dataReader(1))
        Loop
        dataReader.Close()

        Return fields
    End Function

    ''' <summary>
    ''' Determines if a field is contained in a list of fields, searching by field name and field type
    ''' </summary>
    ''' <param name="fieldName">Field name to find</param>
    ''' <param name="dataType">Data Type of field</param>
    ''' <param name="fields">List of fields</param>
    ''' <returns>True if found, false if not</returns>
    Shared Function findFieldByNameAndType(ByVal fieldName As String, ByVal dataType As String, ByVal fields As Dictionary(Of String, String)) As Boolean
        findFieldByNameAndType = False
        For Each campo In fields
            If campo.Key.ToUpper = fieldName.ToUpper And dataType = campo.Value Then _
                Return True
        Next
    End Function

    ''' <summary>
    ''' Gets the list of fields forming the primary key of a table
    ''' </summary>
    ''' <param name="connection">Connection to use</param>
    ''' <param name="tableName">Table to get primary key from</param>
    ''' <returns>Dictionary of fields forming the primary key</returns>
    Shared Function getPrimaryKey(connection As OdbcConnection, tableName As String) As Dictionary(Of String, String)
        Dim fields As New Dictionary(Of String, String)
        Dim query As String = String.Format(" SELECT COLUMN_NAME, TABLE_SCHEMA FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE CONSTRAINT_NAME IN (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS WHERE TABLE_NAME = '{0}' AND CONSTRAINT_TYPE = 'PRIMARY KEY');", tableName)
        Dim cmd As New OdbcCommand(query, connection)
        Dim dataReader As OdbcDataReader = cmd.ExecuteReader()
        Do While dataReader.Read()
            fields.Add(dataReader(0), dataReader(1))
        Loop
        dataReader.Close()
        Return fields
    End Function

    ''' <summary>
    ''' Compares the schemas of two tables residing in two diffrent connections. Checks for same amount of fields, same names, same types and same primary key
    ''' </summary>
    ''' <param name="cnx1">Connection of first table</param>
    ''' <param name="table1Name">First table's name</param>
    ''' <param name="cnx2">Connection of second table</param>
    ''' <param name="table2Name">Second table's name</param>
    ''' <returns>True if all conditions are met, false if at least one fails</returns>
    Shared Function compareSchemas(cnx1 As OdbcConnection, table1Name As String, cnx2 As OdbcConnection, table2Name As String) As Boolean
        compareSchemas = False

        Dim table1Schema As Dictionary(Of String, String) = getSchema(cnx1, table1Name)
        Dim table2Schema As Dictionary(Of String, String) = getSchema(cnx2, table2Name)
        Dim table1PrimaryKey As Dictionary(Of String, String) = getPrimaryKey(cnx1, table1Name)
        Dim table2PrimaryKey As Dictionary(Of String, String) = getPrimaryKey(cnx2, table2Name)

        If table1Schema.Count <> table2Schema.Count OrElse table2PrimaryKey.Count <> table1PrimaryKey.Count Then _
            Return False

        For Each field In table2Schema
            If findFieldByNameAndType(field.Key, field.Value, table1Schema) = False Then _
                Return False
        Next
        For Each campo In table2PrimaryKey
            If findFieldByNameAndType(campo.Key, campo.Value, table1PrimaryKey) = False Then _
                Return False
        Next
        Return True
    End Function

    ''' <summary>
    ''' Get an image stored in the database as an Image object
    ''' </summary>
    ''' <param name="tableName">Table containing the image</param>
    ''' <param name="field">Field containing the image</param>
    ''' <param name="idField">Field identifying the record</param>
    ''' <param name="idValue">Value identifying the record</param>
    ''' <param name="connection">ODBC Connection</param>
    ''' <returns>The image as an Image Object</returns>
    Public Shared Function getImage(tableName As String, field As String, idField As String, idValue As String, connection As OdbcConnection) As Image
        Dim imageByteArray() As Byte = Nothing
        Dim oMemoryStream As MemoryStream = Nothing
        Dim bmpImage As Bitmap = Nothing

        Dim query As String = String.Format("SELECT {1} AS img FROM {0} WHERE {2}={3} AND {2} IS NOT NULL", tableName, field, idField, idValue)

        Dim cmd As New OdbcCommand(query, connection)
        Dim dataReader As OdbcDataReader = cmd.ExecuteReader()

        If dataReader.Read Then
            If Not IsDBNull(dataReader!img) Then
                imageByteArray = CType(dataReader!img, Byte())
                oMemoryStream = New MemoryStream(imageByteArray)
                bmpImage = New Bitmap(oMemoryStream)
            End If
        End If
        dataReader.Close()

        Return bmpImage
    End Function

    ''' <summary>
    ''' Get the data type for a column in a table
    ''' </summary>
    ''' <param name="tableName">Table that contains the column</param>
    ''' <param name="columnName">Column Name</param>
    ''' <param name="connection">ODBC Connection to use</param>
    ''' <returns>The name of the data type</returns>
    Shared Function dataType(ByVal tableName As String, ByVal columnName As String, Optional connection As OdbcConnection = Nothing) As String
        dataType = "not defined"

        Dim sql As String = String.Format("EXEC sp_columns @table_name ='{0}',@column_name='{1}'", tableName, columnName)
        Dim dr As OdbcDataReader = vnframework.ExecuteReader(sql, connection)
        If dr.Read Then
            If dr!TYPE_NAME.ToString.ToLower = "bit" Or dr!TYPE_NAME.ToString.ToLower = "tinyint" Or dr!TYPE_NAME.ToString.ToLower = "smallint" Or dr!TYPE_NAME.ToString.ToLower = "int" Or dr!TYPE_NAME.ToString.ToLower = "bigint" Then
                dataType = "int"
            ElseIf dr!TYPE_NAME.ToString.ToLower = "decimal" Or dr!TYPE_NAME.ToString.ToLower = "float" Or dr!TYPE_NAME.ToString.ToLower = "real" Then
                dataType = "decimal"
            ElseIf dr!TYPE_NAME.ToString.ToLower = "money" Or dr!TYPE_NAME.ToString.ToLower = "smallmoney" Then
                dataType = "money"
            ElseIf dr!TYPE_NAME.ToString.ToLower = "datetime" Or dr!TYPE_NAME.ToString.ToLower = "smalldatetime" Or dr!TYPE_NAME.ToString.ToLower = "timestamp" Then
                dataType = "datetime"
            ElseIf dr!TYPE_NAME.ToString.ToLower = "char" Or dr!TYPE_NAME.ToString.ToLower = "varchar" Or dr!TYPE_NAME.ToString.ToLower = "nchar" Or dr!TYPE_NAME.ToString.ToLower = "nvarchar" Or dr!TYPE_NAME.ToString.ToLower = "nvarchar" Then
                dataType = "varchar"
            ElseIf dr!TYPE_NAME.ToString.ToLower = "binary" Or dr!TYPE_NAME.ToString.ToLower = "varbinary" Or dr!TYPE_NAME.ToString.ToLower = "varbinary" Then
                dataType = "binary"
            Else
                dataType = "not defined"
            End If
        End If
        dr.Close()
    End Function

    ''' <summary>
    ''' Determines if a column accepts null
    ''' </summary>
    ''' <param name="tableName">Table that contains the column</param>
    ''' <param name="columnName">Column Name</param>
    ''' <param name="connection">ODBC Connection to use</param>
    ''' <returns>True if it accepts null; false if it does not</returns>
    Shared Function acceptsNull(ByVal tableName As String, ByVal columnName As String, Optional connection As OdbcConnection = Nothing) As Boolean
        acceptsNull = False
        Dim sql As String = String.Format("EXEC sp_columns @table_name ='{0}', @column_name='{1}'", tableName, columnName)
        Dim dr As OdbcDataReader = vnframework.ExecuteReader(sql, connection)
        If dr.Read Then acceptsNull = IIf(dr!IS_NULLABLE = "YES", True, False)
        dr.Close()
    End Function

    ''' <summary>
    ''' Sets the password for a SQL Server level user
    ''' </summary>
    ''' <param name="user">Username to set password for</param>
    ''' <param name="newPassword">Password to set</param>
    ''' <param name="connection">SQL Connection to use</param>
    ''' <returns>True if action was successful, false if it was not</returns>
    Shared Function setSQLUserPassword(ByVal user As String, ByVal newPassword As String, connection As SqlConnection) As Boolean
        Try
            Dim SQL As String = "EXEC sys.sp_password @New='" & Trim(newPassword) & "', @loginame = '" & Trim(user) & "';"
            Dim SQLCommand As New SqlCommand(SQL, connection)
            SQLCommand.ExecuteNonQuery()
            SQLCommand.Dispose()
            Return True
        Catch ex As SqlException
            logs.writeToErrorLog(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determines if a value exists in a field, excluding a particular record
    ''' </summary>
    ''' <param name="tableName">Table to search in</param>
    ''' <param name="columnName">field to search in</param>
    ''' <param name="value">Value to search for</param>
    ''' <param name="primaryKey">Primary Key name</param>
    ''' <param name="pkValue">Primary Key Value to exclude</param>
    ''' <param name="cnx">Connection to use</param>
    ''' <returns>True if it exists, false if it does not</returns>
    Shared Function valueExists(ByVal tableName As String, ByVal columnName As String, value As String, primaryKey As String, pkValue As String, Optional cnx As OdbcConnection = Nothing) As Boolean
        valueExists = False
        Dim sql As String = String.Format("SELECT COUNT(*) FROM {0} WHERE {1} LIKE '{2}' AND {3} <> '{4}'", tableName, columnName, value, primaryKey, pkValue)
        valueExists = (CInt(ExecuteScalar(sql, cnx)) > 0)
    End Function

    ''' <summary>
    ''' Executes a SQL Server Backup
    ''' </summary>
    ''' <param name="path">Path to sore backup at</param>
    ''' <param name="connection">Connection to use</param>
    ''' <param name="timeout">Timeout in minutes</param>
    ''' <returns>True if successful, false if it fails</returns>
    Shared Function backup(ByVal path As String, connection As OleDbConnection, Optional timeout As Integer = 15) As Boolean
        Using connection
            Try
                Dim sShrink As String = String.Format("DUMP TRAN {0} WITH TRUNCATE_ONLY", connection.Database)
                Dim sBackup As String = String.Format("BACKUP DATABASE {0} TO DISK = N'{1}' WITH NOFORMAT, INIT, NAME =N'{0}-Full Database Backup', SKIP, STATS=10", connection.Database, path)
                Using cmdBackUp As New OleDbCommand(sShrink, connection)
                    Try
                        cmdBackUp.ExecuteNonQuery()
                    Catch ex2 As Exception
                        'Try to shrink, but don't mind if it does not
                    End Try
                    cmdBackUp.CommandText = sBackup
                    cmdBackUp.CommandTimeout = timeout * 60
                    cmdBackUp.ExecuteNonQuery()
                End Using
                connection.Close()
                Return True
            Catch ex As Exception
                showError(ex.Message)
                Return False
            End Try
        End Using
    End Function

    ''' <summary>
    ''' Load a file to a field in a table
    ''' </summary>
    ''' <param name="path">Path to file</param>
    ''' <param name="table">Name of table</param>
    ''' <param name="field">Name of field</param>
    ''' <param name="idField">Primary Key of Table</param>
    ''' <param name="idValue">Value of primary key</param>
    ''' <param name="connection">Connection to use</param>
    ''' <returns>True if successful; false if it fails</returns>
    Shared Function uploadFile(path As String, table As String, field As String, idField As String, idValue As String, connection As SqlConnection) As Boolean
        uploadFile = False
        Try
            Dim sql As String = String.Format("UPDATE {0} SET {1}=@archivo WHERE {2}={3}", table, field, idField, idValue)
            Dim ms As MemoryStream = New MemoryStream()
            Dim fs As FileStream = New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)

            ms.SetLength(fs.Length)
            fs.Read(ms.GetBuffer(), 0, fs.Length)

            Dim arrImg() As Byte = ms.GetBuffer()
            ms.Flush()
            fs.Close()

            Using connection
                Using cmd As SqlCommand = connection.CreateCommand()
                    If connection.State <> ConnectionState.Open Then connection.Open()
                    cmd.CommandText = sql
                    cmd.Parameters.Add("@archivo", SqlDbType.VarBinary).Value = arrImg

                    cmd.ExecuteNonQuery()
                    uploadFile = True
                End Using
            End Using
            ms.Close()
            fs.Dispose()
            ms.Dispose()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Loads a file's contents to a table
    ''' </summary>
    ''' <param name="fileName">Name of file</param>
    ''' <param name="tableName">Table Name</param>
    ''' <param name="fields">Fields to load</param>
    ''' <param name="condition">Connection to use</param>
    ''' <returns>True if successful; false if it fails</returns>
    Public Shared Function textFileToTable(fileName As String, tableName As String, fields As List(Of String), Optional condition As String = "1=1") As Boolean
        textFileToTable = False
        Try
            Dim scampoid As String = ""
            Dim scampos As String = ""
            Dim scamposUpd As String = ""
            Dim i As Integer = 0
            For Each campo As String In fields
                scampos &= IIf(scampos <> "", ",", "") & campo
                If i = 0 Then
                    scampoid = campo
                Else
                    scamposUpd &= IIf(scamposUpd <> "", ",", "") & campo & "= b." & campo
                End If
                i = 1
            Next

            Dim sql_arreglo As New List(Of String)
            sql_arreglo.Add(String.Format("SELECT * INTO {0}_temp FROM {1} WHERE 1=2", "#" & tableName, tableName))
            sql_arreglo.Add(String.Format("BULK INSERT {0}_temp FROM '{1}' WITH (FIELDTERMINATOR = '\t')", "#" & tableName, fileName))
            sql_arreglo.Add(String.Format("INSERT INTO {0} ({1}) SELECT {1} FROM {4}_temp a WHERE (SELECT COUNT(*) FROM {0} b WHERE a.{2}=b.{2}) = 0 AND {3}", tableName, scampos, scampoid, condition, "#" & tableName))
            sql_arreglo.Add(String.Format("UPDATE {0} SET {1} FROM {0} a, {3}_temp b WHERE a.{2}=b.{2}", tableName, scamposUpd, scampoid, "#" & tableName))
            sql_arreglo.Add(String.Format("DROP TABLE {0}_temp", "#" & tableName))

            If procedure(sql_arreglo) = 1 Then _
                textFileToTable = True
        Catch ex As Exception
            showError(ex.Message)
            logs.writeToErrorLog(ex.Message)
        End Try
    End Function
End Class