Imports System.Windows.Forms
Imports System.Data.Odbc
Imports vnframework

Public Class TreeViewEx
    Inherits TreeView

    Private _sql As String
    Property cnx As OdbcConnection

    Property fillQuery As String
        Get
            Return _sql
        End Get
        Set(value As String)
            If Not value Is Nothing Then If value.Length > 0 Then populate(value)
            _sql = value
        End Set
    End Property

    Private Sub populate(ByVal query As String)
        Me.Nodes.Clear()

        Dim dataSetArbol As System.Data.DataSet
        Dim tablaArbol As DataTable

        dataSetArbol = New DataSet("DataSetArbol")

        tablaArbol = dataSetArbol.Tables.Add("TablaArbol")
        tablaArbol.Columns.Add("NombreNodo", GetType(String))
        tablaArbol.Columns.Add("IdentificadorNodo", GetType(String))
        tablaArbol.Columns.Add("IdentificadorPadre", GetType(String))
        tablaArbol.Columns.Add("Seleccionado", GetType(Integer))

        Dim nuevaFila As DataRow
        If _cnx Is Nothing Then

        End If
        Dim dr As OdbcDataReader = ExecuteReader(query, _cnx)

        While dr.Read
            nuevaFila = dataSetArbol.Tables("TablaArbol").NewRow()
            nuevaFila("NombreNodo") = dr(1)
            nuevaFila("IdentificadorNodo") = dr(0)
            nuevaFila("IdentificadorPadre") = sql.Nz(dr(2), 0)
            nuevaFila("Seleccionado") = 0
            If dr.FieldCount >= 5 Then
                CheckBoxes = True
                Try
                    nuevaFila("Seleccionado") = sql.Nz(dr("seleccionado"), 0)
                Catch
                End Try
            End If
            dataSetArbol.Tables("TablaArbol").Rows.Add(nuevaFila)
        End While
        dr.Close()

        populateNodes(0, Nothing, dataSetArbol)

        Me.ExpandAll()
        dataSetArbol.Dispose()
    End Sub

    Public Property insertQuery As String
    Public Function getSaveArrayList() As ArrayList
        Dim sql As New ArrayList
        Dim i As Integer = 0
        For Each node As TreeNode In GetAllNodes()
            If node.Checked Then
                Dim query As String = insertQuery
                If (insertQuery.IndexOf("{parent}") >= 0 Or insertQuery.IndexOf("{parent-R}") >= 0) And node.Parent Is Nothing Then Continue For
                query = query.Replace("{i}", i)
                query = query.Replace("{this}", node.Name)
                If node.Parent Is Nothing Then
                    query = query.Replace("{parent}", "NULL")
                    query = query.Replace("{parent-R}", "NULL")
                Else
                    query = query.Replace("{parent}", node.Parent.Name)
                    query = query.Replace("{parent-R}", node.Parent.Name.Replace("R", ""))
                End If
                i += 1
                sql.Add(query)
            End If
        Next
        Return sql
    End Function

    Private Sub populateNodes(ByVal indicePadre As String, ByVal nodePadre As TreeNode, dataSetArbol As System.Data.DataSet)

        Dim dataViewHijos As DataView

        ' Crear un DataView con los Nodos que dependen del Nodo padre pasado como parámetro.
        dataViewHijos = New DataView(dataSetArbol.Tables("TablaArbol"))

        dataViewHijos.RowFilter = dataSetArbol.Tables("TablaArbol").Columns("IdentificadorPadre").ColumnName + " = '" + indicePadre.ToString() + "'"

        ' Agregar al TreeView los nodos Hijos que se han obtenido en el DataView.
        For Each dataRowCurrent As DataRowView In dataViewHijos

            Dim nuevoNodo As New TreeNode
            nuevoNodo.Text = dataRowCurrent("NombreNodo").ToString().Trim()
            nuevoNodo.Name = dataRowCurrent("IdentificadorNodo").ToString().Trim()
            If CheckBoxes Then nuevoNodo.Checked = CBool(dataRowCurrent("Seleccionado"))

            ' si el parámetro nodoPadre es nulo es porque es la primera llamada, son los Nodos
            ' del primer nivel que no dependen de otro nodo.
            If nodePadre Is Nothing Then
                Me.Nodes.Add(nuevoNodo)
            Else
                ' se añade el nuevo nodo al nodo padre.
                nodePadre.Nodes.Add(nuevoNodo)
            End If

            ' Llamada recurrente al mismo método para agregar los Hijos del Nodo recién agregado.
            populateNodes((dataRowCurrent("IdentificadorNodo").ToString()), nuevoNodo, dataSetArbol)
        Next dataRowCurrent
        dataViewHijos.Dispose()

    End Sub

    Private Function RecurseNodes(ByVal col As TreeNodeCollection) As ArrayList
        RecurseNodes = New ArrayList
        For Each tn As TreeNode In col
            RecurseNodes.Add(tn)
            If tn.Nodes.Count > 0 Then RecurseNodes.AddRange(RecurseNodes(tn.Nodes))
        Next tn
    End Function

    Public Function GetAllNodes() As ArrayList
        GetAllNodes = RecurseNodes(Me.Nodes)
    End Function

End Class