Imports System.Data.OleDb

''' <summary>
''' Function to convert a data type name to objects
''' </summary>
Public Class dataTypes
    ''' <summary>
    ''' Gets the DBType for common field types
    ''' </summary>
    ''' <param name="typeName">Name of field type</param>
    ''' <returns>DBType representing the typeName</returns>
    Shared Function toDBType(typeName As String) As DbType
        typeName = Trim(typeName.Replace("identity", ""))
        Select Case typeName
            Case "int", "integer", "smallint"
                Return DbType.Int64
            Case "varchar"
                Return DbType.AnsiString
            Case "nvarchar"
                Return DbType.AnsiString
            Case "bit"
                Return DbType.Boolean
            Case "ntext", "text"
                Return DbType.AnsiString
            Case "datetime", "date"
                Return DbType.Date
            Case "decimal"
                Return DbType.Decimal
            Case "float"
                Return DbType.Double
            Case "real"
                Return DbType.Double
            Case "numeric"
                Return DbType.Decimal
            Case Else
                Throw New Exception(String.Format(ERROR_TYPE_NOT_DEFINED, typeName))
        End Select
    End Function
    ''' <summary>
    ''' Gets the OleDBType for common field types
    ''' </summary>
    ''' <param name="typeName">Name of field type</param>
    ''' <returns>OleDBType representing the typeName</returns>
    Shared Function toOleDBType(typeName As String) As OleDbType
        typeName = Trim(typeName.Replace("identity", ""))
        Select Case typeName
            Case "int", "integer", "smallint"
                Return OleDbType.Integer
            Case "nvarchar", "varchar", "ntext", "text"
                Return OleDbType.VarChar
            Case "bit"
                Return OleDbType.Boolean
            Case "datetime", "date"
                Return OleDbType.Date
            Case "decimal"
                Return OleDbType.Decimal
            Case "float", "real"
                Return OleDbType.Double
            Case "numeric"
                Return OleDbType.Decimal
            Case Else
                Throw New Exception(String.Format(ERROR_TYPE_NOT_DEFINED, typeName))
        End Select
    End Function
End Class
