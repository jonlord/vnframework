Imports System.Windows.Forms
Public Class TabControlEx
    Public myErrorProvider As ErrorProvider
    Sub switchToFirstTabWithError()
        If myErrorProvider Is Nothing Then Exit Sub
        For Each myTabPage As TabPage In TabPages
            For Each ctrl As Control In myTabPage.Controls
                If myErrorProvider.GetError(ctrl) <> "" Then
                    SelectTab(myTabPage.Name)
                    Exit Sub
                End If
                On Error Resume Next
            Next
        Next
    End Sub
End Class