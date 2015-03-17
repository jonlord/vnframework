Imports System.ComponentModel

''' <summary>
''' Enumerator helper functions
''' </summary>
Public Class enums
    ''' <summary>
    ''' Gets an enumerator's description from its constant
    ''' </summary>
    ''' <param name="enumConstant">The enumerators constant value</param>
    ''' <returns>The enumerators description</returns>
    Public Shared Function enumDescription(ByVal enumConstant As [Enum]) As String
        If enumConstant Is Nothing Then Return Nothing
        Dim fi As Reflection.FieldInfo = enumConstant.GetType().GetField(enumConstant.ToString())
        Dim aattr() As DescriptionAttribute = DirectCast(fi.GetCustomAttributes(GetType(DescriptionAttribute), False), DescriptionAttribute())
        If aattr.Length > 0 Then
            Return aattr(0).Description
        Else
            Return enumConstant.ToString()
        End If
    End Function
    ''' <summary>c
    ''' Gets an enumerator's value from its constant
    ''' </summary>
    ''' <param name="enumConstant">The enumerators constant value</param>
    ''' <returns>The enumerators value</returns>
    Public Shared Function enumValue(ByVal enumConstant As [Enum]) As String
        If enumConstant Is Nothing Then Return Nothing
        Dim fi As Reflection.FieldInfo = enumConstant.GetType().GetField(enumConstant.ToString())
        Dim aattr() As AmbientValueAttribute = DirectCast(fi.GetCustomAttributes(GetType(AmbientValueAttribute), False), AmbientValueAttribute())
        If aattr.Length > 0 Then
            Return aattr(0).Value
        Else
            Return enumConstant.ToString()
        End If
    End Function
End Class