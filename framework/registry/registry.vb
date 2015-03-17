Imports Microsoft.Win32

''' <summary>
''' Windows Registry manipulation library
''' </summary>
Public Class windowsRegistry
    Private SubKeyName As String

    ''' <summary>
    ''' Initiates a registry object selecting a subkey to use
    ''' </summary>
    ''' <param name="subKeyName">SubKey to use</param>
    Sub New(Optional ByVal subKeyName As String = "SOFTWARE\")
        setSubKey(subKeyName)
    End Sub

    ''' <summary>
    ''' Sets the subkey to use
    ''' </summary>
    ''' <param name="subKeyName">SubKey Name</param>
    Public Sub setSubKey(Optional ByVal subKeyName As String = "SOFTWARE\")
        Me.SubKeyName = subKeyName
    End Sub
    ''' <summary>
    ''' Writes a value to the registry
    ''' </summary>
    ''' <param name="subKey">Registry Subkey in Hive</param>
    ''' <param name="value">Value to write</param>
    ''' <returns>True if successfull or false if an error occurs</returns>
    Public Function WriteToRegistry(ByVal subKey As String, ByVal value As Object) As Boolean
        Dim ParentKeyHive As RegistryHive = RegistryHive.CurrentUser
        Dim objSubKey As RegistryKey
        Dim objParentKey As RegistryKey = Nothing
        Try
            Select Case ParentKeyHive
                Case RegistryHive.ClassesRoot
                    objParentKey = Registry.ClassesRoot
                Case RegistryHive.CurrentConfig
                    objParentKey = Registry.CurrentConfig
                Case RegistryHive.CurrentUser
                    objParentKey = Registry.CurrentUser
                Case RegistryHive.LocalMachine
                    objParentKey = Registry.LocalMachine
                Case RegistryHive.PerformanceData
                    objParentKey = Registry.PerformanceData
                Case RegistryHive.Users
                    objParentKey = Registry.Users
            End Select
            Dim array As String() = subKey.Split("\")
            Dim ext As String = ""
            If array.Length > 1 Then
                For i As Integer = 0 To array.Length - 2
                    ext = "\" & array(i)
                Next i
                subKey = array(array.Length - 1)
            End If
            objSubKey = objParentKey.OpenSubKey(SubKeyName & ext, True)
            If objSubKey Is Nothing Then objSubKey = objParentKey.CreateSubKey(SubKeyName & ext)
            objSubKey.SetValue(subKey, value)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    ''' <summary>
    ''' Get Value from registry
    ''' </summary>
    ''' <param name="subKey">Registry Subkey in Hive</param>
    ''' <param name="defaultValue">Value to return if the subkey is not found or an error occurs</param>
    ''' <returns>Value of SubKey of defaultValue if error or nor found</returns>
    Public Function GetFromRegistry(ByVal subKey As String, Optional ByRef defaultValue As String = "") As String
        Dim Hive As RegistryHive = RegistryHive.CurrentUser
        Dim objParent As RegistryKey = Nothing
        Dim objSubkey As RegistryKey
        Select Case Hive
            Case RegistryHive.ClassesRoot
                objParent = Registry.ClassesRoot
            Case RegistryHive.CurrentConfig
                objParent = Registry.CurrentConfig
            Case RegistryHive.CurrentUser
                objParent = Registry.CurrentUser
            Case RegistryHive.LocalMachine
                objParent = Registry.LocalMachine
            Case RegistryHive.PerformanceData
                objParent = Registry.PerformanceData
            Case RegistryHive.Users
                objParent = Registry.Users
        End Select
        Try
            Dim array As String() = subKey.Split("\")
            Dim ext As String = ""
            If array.Length > 1 Then
                For i As Integer = 0 To array.Length - 2
                    ext = "\" & array(i)
                Next i
                subKey = array(array.Length - 1)
            End If
            objSubkey = objParent.OpenSubKey(SubKeyName & ext)
            If Not objSubkey Is Nothing Then
                Dim val As String = objSubkey.GetValue(subKey)
                If val Is Nothing Then Return defaultValue
                Return val
            End If
        Catch ex As Exception
            Return defaultValue
        End Try
        Return defaultValue
    End Function
End Class