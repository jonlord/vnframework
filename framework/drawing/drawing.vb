Imports System.Windows.Forms
Imports System.Drawing
Public Module drawingEx
    Public Sub DrawLine(ByVal control As Control, ByVal color As Color, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)
        Dim myPen As Pen = New Pen(color, 1)
        Dim bit As Bitmap = New Bitmap(control.Width, control.Height)
        Dim g As Graphics = Graphics.FromImage(bit)
        control.CreateGraphics.DrawLine(myPen, x1, y1, x2, y2)
    End Sub
    Public Sub DrawRectangle(ByVal control As Control, ByVal color As Color, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)
        Dim myPen As New System.Drawing.Pen(color)
        Dim formGraphics As System.Drawing.Graphics
        formGraphics = control.CreateGraphics()
        formGraphics.DrawRectangle(myPen, New Rectangle(x1, y1, x2, y2))
        myPen.Dispose()
        formGraphics.Dispose()
    End Sub
    Public Sub DrawFilledRectangle(ByVal control As Control, ByVal color As Color, ByVal fillColor As Color, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)
        Dim myBrush As New SolidBrush(fillColor)
        Dim formGraphics As System.Drawing.Graphics
        formGraphics = control.CreateGraphics()
        formGraphics.FillRectangle(myBrush, New Rectangle(x1, y1, x2, y2))
        DrawRectangle(control, color, x1, y1, x2, y2)
        myBrush.Dispose()
        formGraphics.Dispose()
    End Sub
    Public Sub DrawFourthCircle(ByVal control As Control, ByVal color As Color, ByVal fillColor As Color, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal a1 As Integer, ByVal a2 As Integer)
        Dim myBrushOuter As New System.Drawing.SolidBrush(color)
        Dim myBrushInner As New System.Drawing.SolidBrush(fillColor)
        Dim formGraphics As System.Drawing.Graphics
        formGraphics = control.CreateGraphics()
        If a1 >= 270 Then
            formGraphics.FillPie(myBrushOuter, New Rectangle(x1 + 1, y1, x2 + 1, y2), a1, a2)
            formGraphics.FillPie(myBrushInner, New Rectangle(x1 + 1, y1 + 1, x2, y2 - 1), a1, a2)
        Else
            formGraphics.FillPie(myBrushOuter, New Rectangle(x1, y1, x2, y2), a1, a2)
            formGraphics.FillPie(myBrushInner, New Rectangle(x1 + 1, y1 + 1, x2 - 1, y2 - 1), a1, a2)
        End If
        myBrushOuter.Dispose()
        myBrushInner.Dispose()
        formGraphics.Dispose()
    End Sub
    Public Sub DrawGradient(ByVal control As Control, ByVal color1 As Color, ByVal color2 As Color, ByVal mode As System.Drawing.Drawing2D.LinearGradientMode)
        Dim a As New System.Drawing.Drawing2D.LinearGradientBrush(New RectangleF(0, 0, control.Width, control.Height), color1, color2, mode)
        Dim g As Graphics = control.CreateGraphics
        g.FillRectangle(a, New RectangleF(0, 0, control.Width, control.Height))
        g.Dispose()
    End Sub
    Public Sub DrawGradientString(ByVal control As Control, ByVal text As String, ByVal color1 As Color, ByVal color2 As Color, ByVal mode As System.Drawing.Drawing2D.LinearGradientMode)
        Dim a As New System.Drawing.Drawing2D.LinearGradientBrush(New RectangleF(0, 0, 100, 19), color1, color2, mode)
        Dim g As Graphics = control.CreateGraphics
        Dim f As Font
        f = New Font("arial", 20, FontStyle.Bold, GraphicsUnit.Pixel)
        g.DrawString(text, f, a, 0, 0)
        g.Dispose()
    End Sub
End Module