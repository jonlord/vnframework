Imports System.ServiceProcess

''' <summary>
''' Collection of non-original functions to manage machine startup up and shutdown
''' </summary>
Public Class WindowsOperations

    ''' <summary>
    ''' Start a windows service
    ''' </summary>
    ''' <param name="server">FQDN or IP of server to control</param>
    ''' <param name="serviceName">Name of service to control</param>
    ''' <returns>True if able to start; false if not</returns>
    Shared Function startService(server As String, serviceName As String) As Boolean
        Dim status As String  'Service status
        Using mySC As ServiceController = New ServiceController(serviceName, server)
            Try
                'Save service status
                status = mySC.Status.ToString
                If status = "" Then Return False
            Catch ex As Exception
                showError(String.Format(ERRORSERVICENOTFOUND, ex.Message))
                Return False
            End Try
            'If the service is stopped, try to start it
            If mySC.Status.Equals(ServiceControllerStatus.Stopped) Or mySC.Status.Equals(ServiceControllerStatus.StopPending) Then
                Try
                    mySC.Start()
                    'Wait for service to start
                    mySC.WaitForStatus(ServiceControllerStatus.Running)
                Catch ex As Exception
                    'Report an error if not able to start
                    showError(String.Format(ERRORSTARTINGSERVICE, ex.Message))
                    Return False
                End Try
            End If
        End Using
        Return True
    End Function

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

    ''' <summary>
    ''' Gets the path of the current application
    ''' </summary>
    ''' <returns>The path to the directory where the application resides</returns>
    Public Function appPath() As String
        Return AppDomain.CurrentDomain.BaseDirectory()
    End Function

    Private Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As Long, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean

    ''' <summary>
    ''' Get the path to a special folder
    ''' </summary>
    ''' <param name="folderName">Special folder name</param>
    ''' <returns>Path to special folder</returns>
    Public Function getSpecialFolderA(ByVal folderName As Environment.SpecialFolder) As String
        Dim typeSpecialFolder As Environment.SpecialFolder
        For Each typeSpecialFolder In typeSpecialFolder.GetValues(GetType(Environment.SpecialFolder))
            If typeSpecialFolder = folderName Then
                Return Environment.GetFolderPath(typeSpecialFolder)
            End If
        Next
        Return ""
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