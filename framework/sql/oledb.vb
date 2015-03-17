Imports System.Data.OleDb
''' <summary>
''' OleDB Manipulation Library
''' </summary>
Public Class oledb
    ''' <summary>
    ''' Gets the OleDB equivalent for a System.Type
    ''' </summary>
    ''' <param name="sysType">System.Type to convert</param>
    ''' <param name="fieldName">Reference usage</param>
    ''' <returns>The OleDB equivalent for a System.Type</returns>
    Public Shared Function GetOleDbType(ByVal sysType As Type, Optional ByVal fieldName As String = "") As OleDbType
        If sysType Is GetType(String) Then
            Return OleDbType.VarChar
        ElseIf sysType Is GetType(Integer) Or sysType.FullName = "System.UInt16" Or sysType Is GetType(Int16) Or sysType.FullName = "System.UInt32" Or sysType.FullName = "System.UInt64" Or sysType Is GetType(Int32) Or sysType Is GetType(Int32) Then
            Return OleDbType.Integer
        ElseIf sysType.FullName = "System.Int64" Then
            Return OleDbType.BigInt
        ElseIf sysType Is GetType(Boolean) Then
            Return OleDbType.Integer
        ElseIf sysType Is GetType(Date) Then
            Return OleDbType.Date
        ElseIf sysType Is GetType(Char) Then
            Return OleDbType.Char
        ElseIf sysType Is GetType(Decimal) Then
            Return OleDbType.Decimal
        ElseIf sysType Is GetType(Double) Then
            Return OleDbType.Double
        ElseIf sysType Is GetType(Single) Then
            Return OleDbType.Single
        ElseIf sysType Is GetType(Byte()) Then
            Return OleDbType.Binary
        ElseIf sysType Is GetType(Guid) Then
            Return OleDbType.Guid
        ElseIf sysType Is GetType(Byte) Or sysType.FullName = "System.SByte" Then
            Return OleDbType.TinyInt
        Else
            Throw New ArgumentException(String.Format(ERRROOLEDBTYPENOTFOUND, sysType.FullName, fieldName))
        End If
    End Function
End Class