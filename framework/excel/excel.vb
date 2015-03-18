Imports System.Data.OleDb
Imports System.IO

''' <summary>
''' Excel manipulation library
''' </summary>
Public Class excel
    ''' <summary>
    ''' Creates an Excel File from a SQL Query
    ''' </summary>
    ''' <param name="connectionString">Connection string to use</param>
    ''' <param name="SQL">SQL Query to execute</param>
    ''' <param name="path">Path to save excel file to</param>
    ''' <param name="sheetName">Name of sheet to save contents</param>
    Shared Sub sqlToExcel(ByVal connectionString As String, ByVal SQL As String, ByVal path As String, sheetName As String)
        SQL = String.Format("INSERT INTO OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0', 'Data Source={0};Extended Properties=Excel 8.0') ... [{2}$]{1}", path, SQL, sheetName)
        Dim myConnection As New OleDbConnection(connectionString)
        myConnection.Open()
        Dim o As New OleDbCommand(SQL, myConnection)
        o.ExecuteNonQuery()
        myConnection.Close()
    End Sub

    ''' <summary>
    ''' Cleans a cell's contents so its compatible with SQL
    ''' </summary>
    ''' <param name="Value">Current Value</param>
    ''' <param name="valueIfEmpty">Value to replace if empty</param>
    ''' <returns>The cleaned up version of value</returns>
    Public Shared Function CellCleanForSQL(Value As String, Optional valueIfEmpty As String = "") As String
        Value = sql.Nz(Value, valueIfEmpty)
        If Value Is Nothing Then Value = ""
        Return Value.escapeCharacters.Replace(vbNewLine, "").Replace(vbTab, "")
    End Function

    ''' <summary>
    ''' Saves a byte array as an XLSB File
    ''' </summary>
    ''' <param name="array">Array of bytes to save</param>
    ''' <param name="path">Path to save file at</param>
    ''' <returns>True if completed, false if not</returns>
    Shared Function byteArrayToXLSB(ByVal array() As Byte, path As String) As Boolean
        Try
            Dim f As FileStream = File.Create(path)
            For Each b As Byte In array
                f.WriteByte(b)
            Next
            f.Flush()
            f.Close()
            Return True
        Catch ex As Exception
            showError(ex.Message)
            Return False
        End Try
    End Function
End Class