Imports System.Text.RegularExpressions

''' <summary>
''' SQL Strings Manipulation Library
''' </summary>
Public Class sql
    ''' <summary>
    ''' Converts a System.Date to a SQL standard date format
    ''' </summary>
    ''' <param name="theDate">The date to convert</param>
    ''' <param name="includeTime">Include time</param>
    ''' <param name="useLastSecondOfDay">Replace the time with 23:59:59</param>
    ''' <returns>String representation of the ready ready to be used in a SQL Query</returns>
    Public Shared Function convertVBtoDate(ByVal theDate As Date, Optional ByVal includeTime As Boolean = False, Optional ByVal useLastSecondOfDay As Boolean = True) As String
        convertVBtoDate = String.Format("{0}-{1}-{2}", theDate.Year, Right("0" & theDate.Month, 2), Right("0" & theDate.Day, 2))
        If includeTime And useLastSecondOfDay Then convertVBtoDate &= " 23:59:59"
        If includeTime And Not useLastSecondOfDay Then convertVBtoDate &= String.Format(" {0}:{1}:{2}", theDate.Hour, Right("0" & theDate.Minute, 2), Right("0" & theDate.Second, 2))
        convertVBtoDate = String.Format("'{0}'", convertVBtoDate)
    End Function
    ''' <summary>
    ''' Converts a System.Date to a SQL CONVERT function call
    ''' </summary>
    ''' <param name="theDate">The date to convert</param>
    ''' <returns>String representation of the CONVERT function call</returns>
    Public Shared Function convertVarcharToDate(ByVal theDate As String) As String
        Dim Format As String = "120"
        Return String.Format("CONVERT(VARCHAR(10), {0}, {1})", theDate, Format)
    End Function

    '' <summary>
    '' Makes changes to a query to make it compatible with SQL Server
    '' </summary>
    '' <param name="originalQuery">Original query</param>
    '' <returns>SQL query with changes to be compatible with SQL Server</returns>
    Public Shared Function parseToSQLServer(ByVal originalQuery As String) As String
        Try
            originalQuery = Regex.Replace(originalQuery, "\b[Nn][Oo][Ww]\(\)", "GETDATE()")
            originalQuery = Regex.Replace(originalQuery, "\b[Ii][Ff][Nn][Uu][Ll][Ll]\(", "ISNULL(")
            originalQuery = Regex.Replace(originalQuery, "\b[Nn][Zz]\(", "ISNULL(")
            originalQuery = Regex.Replace(originalQuery, "\b[Tt][Rr][Ii][Mm]\((\w+)\)", "RTRIM(LTRIM($1))")
            originalQuery = Regex.Replace(originalQuery, "\b[Ii][Nn][Ss][Tt][Rr]\((\w+),([A-Za-z0-9-'\s]+)\)", "CHARINDEX($2, $1)")
            originalQuery = Regex.Replace(originalQuery, "\b[Ii][Ii][Ff]\(([0-9A-Za-z.<>()_'*=/+ \-]+),([0-9A-Za-z.<>()_'*=/+ \-]+),([0-9A-Za-z.<>()_'*=/+ \-]+)\)", "CASE WHEN $1 THEN $2 ELSE $3 END")
            originalQuery = Regex.Replace(originalQuery, "\b[Mm][Ii][Dd]\(", "SUBSTRING(")
            originalQuery = Regex.Replace(originalQuery, "\b[Aa][Dd][Dd]\s[Cc][Oo][Ll][Uu][Mm][Nn]\s", "ADD ")
        Catch ex As Exception
            Return originalQuery
        End Try
        Return originalQuery
    End Function

    ''' <summary>
    ''' Returns a SQL snippet depending on the database used that functions as a ISNULL / IFNULL
    ''' </summary>
    ''' <param name="value">Value to test if is null</param>
    ''' <param name="valueIfNull">Value to return if the first parameter is null</param>
    ''' <returns>SQL Snippet depending on database model for the ISNULL / IFNULL Function</returns>
    Public Shared Function SQLIfNull(ByVal value As String, ByVal valueIfNull As String) As String
        ' SQLIfNull = String.Format("IIF(ISNULL({0}),{1},{0})", value, valueIfNull) 'For Access
        SQLIfNull = String.Format("ISNULL({0},{1})", value, valueIfNull)
    End Function

    ''' <summary>
    ''' Determines if a value is null and returns another one if so. If it is not null returns the same value
    ''' </summary>
    ''' <param name="valueToTest">Value to test for null</param>
    ''' <param name="returnIfNull">Value to return if null</param>
    ''' <returns>returnIfNull if valueToTest is null, valuetTotTest if not</returns>
    Public Shared Function Nz(ByVal valueToTest As Object, ByVal returnIfNull As Object)
        Try
            If valueToTest Is Nothing Then
                Return returnIfNull
            ElseIf (IsDBNull(valueToTest) OrElse valueToTest.ToString() = "") Then
                Return returnIfNull
            Else
                Return valueToTest
            End If
        Catch e As NullReferenceException
            Return returnIfNull
        End Try
    End Function
End Class