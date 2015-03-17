''' <summary>
''' Collection of non-original functions to manage machine startup up and shutdown
''' </summary>
Public Class WindowsOperations

    ''' <summary>
    ''' Shuts down the computer
    ''' </summary>
    Shared Sub shutdown()
        Dim platform As New PlatformID
        Select Case Environment.OSVersion.Platform
            Case PlatformID.Win32NT
                Dim token As TOKEN_PRIVILEGES
                Dim blank_token As TOKEN_PRIVILEGES
                Dim token_handle As IntPtr
                Dim uid As LUID
                Dim ret_length As Integer
                Dim ptr As IntPtr = GetCurrentProcess()

                OpenProcessToken(ptr, &H20 Or &H8, token_handle)
                LookupPrivilegeValue("", "SeShutdownPrivilege", uid)
                token.PrivilegeCount = 1
                token.Privileges = uid
                token.Attributes = &H2

                AdjustTokenPrivileges(token_handle, 0, token, System.Runtime.InteropServices.Marshal.SizeOf(blank_token), blank_token, ret_length)
                ExitWindowsEx(EWX_SHUTDOWN Or EWX_FORCE, &HFFFF)
            Case Else
                ExitWindowsEx(EWX_SHUTDOWN Or EWX_FORCE, &HFFFF)
        End Select
    End Sub
    ''' <summary>
    ''' Adds a program to the windows registry to make it run upon startup
    ''' </summary>
    ''' <param name="exePath">Executable to run</param>
    ''' <param name="appName">Name to be displayed on the Registry</param>
    ''' <returns>True if successfull, false if it fails</returns>
    Public Shared Function setAutoStart(exePath As String, appName As String) As Boolean
        Try
            Dim CU As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
            With CU
                .OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
                .SetValue(appName, exePath)
            End With
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As IntPtr
    Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As IntPtr, ByVal DesiredAccess As Int32, ByRef TokenHandle As IntPtr) As Int32
    Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LUID) As Int32
    Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As IntPtr, ByVal DisableAllPrivileges As Int32, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Int32, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As Int32) As Int32

    Private Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Int32, ByVal dwReserved As Int32) As Int32
    Private Const EWX_FORCE As Int32 = 4
    Private Const EWX_SHUTDOWN As Int32 = 1
    Private Const EWX_REBOOT As Int32 = 2
    Private Const EWX_LOGOFF As Int32 = 0

    Private Structure LUID
        Public Property LowPart As Int32
        Public Property HighPart As Int32
    End Structure
    Private Structure TOKEN_PRIVILEGES
        Public Property PrivilegeCount As Integer
        Public Property Privileges As LUID
        Public Property Attributes As Int32
    End Structure
End Class