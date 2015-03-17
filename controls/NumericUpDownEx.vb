Imports System.Windows.Forms
Imports System.Drawing

Public Class NumericUpDownEx
    Inherits NumericUpDown

    Protected Overrides Sub OnTextBoxResize(ByVal source As Object, ByVal e As EventArgs)
        Controls(0).Hide()
        Controls(1).Size = New Size(ClientSize.Width - 10, ClientSize.Height - 4)
    End Sub
End Class