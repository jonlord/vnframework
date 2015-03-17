Imports System.IO
Imports System.Runtime.InteropServices

''' <summary>
''' Font helper functions
''' </summary>
Public Class fonts
    Private Const WM_FONTCHANGE As Integer = &H1D
    Private Const HWND_BROADCAST As Integer = &HFFFF

    <DllImport("gdi32")>
    Private Shared Function AddFontResource(ByVal lpFileName As String) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SendMessage(ByVal hWnd As Integer, ByVal Msg As UInteger, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function WriteProfileString(ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Integer
    End Function

    ''' <summary>
    ''' Installs a TTF font to a Window's PC
    ''' </summary>
    ''' <param name="fontName">Font name, (the file's name) including the extension</param>
    ''' <param name="fontPath">Path were the TTF file is stored</param>
    Public Shared Sub addFont(ByVal fontName As String, fontPath As String)
        fontName = fontName.Trim.ToUpper
        If Not fontName.EndsWith("ttf", StringComparison.OrdinalIgnoreCase) Then fontName = fontName + ".ttf"
        Dim theoreticalFontPath As String = String.Format("{0}\{1}", Environment.GetFolderPath(Environment.SpecialFolder.Fonts), fontName)
        If Not File.Exists(theoreticalFontPath) Then
            Dim archfont As Byte() = File.ReadAllBytes(String.Format("{0}/{1}", fontPath, fontName))
            File.WriteAllBytes(theoreticalFontPath.Replace("/", "\"), archfont)

            AddFontResource(theoreticalFontPath.Replace("\", "/"))
            SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0)
        End If
    End Sub
End Class