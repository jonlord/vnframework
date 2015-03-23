Imports vnframework
Imports System.Data.Odbc

''' <summary>
''' Transfer data from one sql server connection to another
''' </summary>
Public Class transferData
    ''' <summary>
    ''' Transfer data from one sql server connection to another
    ''' </summary>
    ''' <param name="tableSource">Source table</param>
    ''' <param name="tableDestiny">Destiny table</param>
    ''' <param name="connectionSource">Source Connection</param>
    ''' <param name="connectionDestiny">Destiny Connection</param>
    ''' <returns>True if completed succesfully, false if not</returns>
    Public Shared Function transferData(ByVal tableSource As String, ByVal tableDestiny As String, connectionSource As OdbcConnection, connectionDestiny As OdbcConnection) As Boolean
        Dim cmdSource As New OdbcCommand
        Dim queries As New ArrayList
        Dim dataReader As OdbcDataReader
        Dim fieldList As Dictionary(Of String, String)
        Dim idFields As String = ""
        Dim fields As String = ""
        Dim identity As String = ""
        Dim ids As New ArrayList

        Try
            connectionDestiny.Open()
        Catch ex As Exception
            showError(ERRORCONNECTDESTINYSERVER & vbNewLine & ex.Message)
            Return False
        End Try

        Try
            connectionSource.Open()
        Catch ex As Exception
            showError(ERRORCONNECTSOURCESERVER & vbNewLine & ex.Message)
            Return False
        End Try

        If sqlserver.compareSchemas(connectionSource, tableSource, connectionDestiny, tableDestiny) Then
            cmdSource.Connection = connectionSource
            fieldList = sqlserver.getSchema(connectionSource, tableSource)
            For Each campo In fieldList
                fields = fields & campo.Key & ","
            Next
            fields = fields.Replace("to", """to""")
            fields = fields.TrimEnd(",")
            identity = ExecuteScalar(String.Format("SELECT ISNULL(name,'') FROM sys.columns WHERE object_id = (SELECT object_id FROM sys.objects WHERE TYPE = 'U' AND name = '{0}') AND is_identity = 1", tableDestiny), connectionDestiny)
            If identity <> "" Then
                fields = fields.Replace(identity & ",", "")
                fields = fields.Replace(identity, "")
                idFields &= identity
            End If
            Dim fieldsToAdd As String = String.Format("{0}{1}", IIf(identity <> "", idFields & ",", ""), fields)
            cmdSource.CommandText = String.Format("SELECT {0} FROM {1};", fieldsToAdd, tableSource)
            Dim pkFuente As Dictionary(Of String, String) = sqlserver.getPrimaryKey(connectionSource, tableSource)

            dataReader = cmdSource.ExecuteReader()
            Do While dataReader.Read()
                Dim values As String = getValues(dataReader, identity, tableDestiny, connectionDestiny, queries)
                queries.Add(String.Format("INSERT INTO {0} ({1}) VALUES ({2})", tableDestiny, fields, values.TrimEnd(",")))

                Dim thePKFields As New ArrayList, thePKValues As New ArrayList
                For Each campo In pkFuente
                    thePKFields.Add(campo.Key)
                    thePKValues.Add(dataReader(campo.Key))
                Next
                ids.Add(New pkValues(thePKFields, thePKValues))
            Loop
            dataReader.Close()
            If queries.Count > 0 Then
                If procedure(queries, , , connectionDestiny) = 1 Then
                    queries.Clear()
                    For Each pkValue As pkValues In ids
                        Dim condition As String = "1=1 "
                        With pkValue
                            For i As Integer = 0 To .fieldName.Count - 1
                                condition &= String.Format(" AND {0}='{1}'", .fieldName(i), .fieldValue(i))
                            Next
                        End With
                        queries.Add(String.Format("DELETE FROM {0} WHERE {1}", tableSource, condition))
                    Next
                    If procedure(queries, , , connectionSource) = 1 Then
                        Return True
                    Else
                        showError(ERRORSAVE)
                        Return False
                    End If
                Else
                    showError(ERRORSAVE)
                    Return False
                End If
            End If
        Else
            showError(ERRORDIFFERENTSCHEMAS)
        End If
        Return True
    End Function

    ''' <summary>
    ''' Get all values of a row a csv string
    ''' </summary>
    ''' <param name="row">ObdcReader positioned on the desired row</param>
    ''' <param name="identity">Identity field</param>
    ''' <param name="table">Table Name</param>
    ''' <param name="cnx">Connection to use</param>
    ''' <param name="queries">List of queries to add in procedure</param>
    ''' <returns>All values of a row a csv string</returns>
    Shared Function getValues(row As OdbcDataReader, identity As String, table As String, cnx As OdbcConnection, ByRef queries As ArrayList) As String
        Dim values As String = ""
        Dim pk As Dictionary(Of String, String) = sqlserver.getPrimaryKey(cnx, table)
        Dim found As Boolean = False
        For i As Integer = 0 To row.FieldCount - 1
            If identity = row.GetName(i) Then
                values = values
            Else
                found = False
                For Each field In pk
                    If field.Key = row.GetName(i) And row(i).GetType.ToString.ToLower = "system.int16" Then
                        found = True
                        Exit For
                    End If
                Next
                If found = False Then
                    Select Case row(i).GetType.ToString.ToLower
                        Case "system.datetime"
                            values = values & "" & sql.convertVBtoDate(row(i), True, False) & ","
                        Case "system.string"
                            values = values & "'" & row(i) & "',"
                        Case "system.int16"
                            values = values & "" & row(i) & ","
                        Case "system.dbnull"
                            values = values & "null,"
                        Case "system.boolean"
                            If row(i) = True Then
                                values = values & "1,"
                            Else
                                values = values & "0,"
                            End If
                        Case Else
                            showError("No se ha especificado el tipo de dato " & row(i).GetType.ToString & ". El error ocurre en la función transferData, cuando se estan armando los values")
                    End Select
                Else
                    queries.Add(String.Format("SELECT ISNULL(MAX(id),0)+1 FROM {0}", table))
                    values = values & String.Format("@@rSQL{0}@@,", queries.Count - 1)
                End If
            End If
        Next
        Return values
    End Function

    Private Class pkValues
        Public fieldName As ArrayList
        Public fieldValue As ArrayList
        Sub New(fieldName As ArrayList, fieldValue As ArrayList)
            Me.fieldName = fieldName
            Me.fieldValue = fieldValue
        End Sub
    End Class
End Class