Imports System.Xml
Imports System.IO
Public Class XmlDocumentEX
    Private xml As XmlDocument
    Sub New(xml As XmlDocument)
        Me.xml = xml
    End Sub
    Public Function toDataTable() As DataTable
        Dim ms As MemoryStream = Nothing
        Dim returnMs As New DataTable()
        Try
            Dim ds As New DataSet
            ds.ReadXml(New XmlNodeReader(xml))
            Return ds.Tables(0)
        Catch ex As Exception
            Return returnMs
        Finally
            If Not ms Is Nothing Then
                ms.Close()
            End If
        End Try
    End Function
End Class

