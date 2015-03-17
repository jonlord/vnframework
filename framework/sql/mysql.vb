''' <summary>
''' MySQL manipulation library
''' </summary>
Public Class mysql
    ''' <summary>
    ''' Gets the MySQL equivalent for a System.Type
    ''' </summary>
    ''' <param name="sysType">System.Type to convert</param>
    ''' <param name="fieldName">Reference usage</param>
    ''' <returns>The OleDB equivalent for a System.Type</returns>
    Public Shared Function GetDbType(ByVal sysType As Type, Optional fieldName As String = "") As DbType
        Dim myMySqlDbType As DbType = DbType.Int64
        If sysType.Equals(GetType(Integer)) Then
            myMySqlDbType = DbType.Int64
        ElseIf sysType.Equals(GetType(Int16)) Then
            myMySqlDbType = DbType.Int16
        ElseIf sysType.Equals(GetType(Int32)) Then
            myMySqlDbType = DbType.Int32
        ElseIf sysType.Equals(GetType(Int64)) Then
            myMySqlDbType = DbType.Int64
        ElseIf sysType.Equals(GetType(Single)) Then
            myMySqlDbType = DbType.Single
        ElseIf (sysType.Equals(GetType(Long)) Or sysType.Equals(GetType(Long))) Then
            myMySqlDbType = DbType.Int64
        ElseIf sysType.Equals(GetType(Double)) Then
            myMySqlDbType = DbType.Double
        ElseIf (sysType.Equals(GetType(Decimal)) Or sysType.Equals(GetType(Decimal))) Then
            myMySqlDbType = DbType.Decimal
        ElseIf (sysType.Equals(GetType(DateTime))) Then
            myMySqlDbType = DbType.DateTime
        ElseIf (sysType.Equals(GetType(Boolean)) Or sysType.Equals(GetType(Byte)) Or sysType.Equals(GetType(Byte))) Then
            myMySqlDbType = DbType.Boolean
        ElseIf (sysType.Equals(GetType(String))) Then
            myMySqlDbType = DbType.String
        ElseIf (sysType.Equals(GetType(Byte())) Or sysType.Equals(GetType(Byte()))) Then
            'hmm.. possible type size loss here as we dont know the length of the array
            myMySqlDbType = DbType.String
        Else
            Throw New ArgumentException(String.Format(ERRROMYSQLTYPENOTFOUND, sysType.FullName, fieldName))
        End If

        Return myMySqlDbType
    End Function
    ''' <summary>
    ''' Fixes values to be compatible with MySQL
    ''' </summary>
    ''' <param name="value">Value to convert</param>
    ''' <param name="valueType">Value type</param>
    ''' <returns>Converted Value</returns>
    Public Shared Function convertValue(ByVal value As Object, ByVal valueType As DbType) As Object
        If valueType <> DbType.Boolean Then Return value
        If CType(value, Boolean) Then Return 1
        Return 0
    End Function
End Class