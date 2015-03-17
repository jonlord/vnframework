''' <summary>
''' Make an audible alarm
''' </summary>
Public Module audibleAlarm
    ''' <summary>
    ''' Message Beep Types
    ''' </summary>
    Public Enum messageBeepType
        [Default] = -1
        OK = &H0
        [Error] = &H10
        Question = &H20
        Warning = &H30
        Information = &H40
    End Enum

    ''' <summary>
    ''' Makes a beep from predefined types
    ''' </summary>
    ''' <param name="messageBeepType">Type of Beep</param>
    Public Sub audibleBeep(messageBeepType As messageBeepType)
        MessageBeep(messageBeepType)
    End Sub

    ''' <summary>
    ''' This allows you to 'beep' using one of the sounds mapped using the' sound mapper in control panel.
    ''' </summary>
    ''' <param name="type">Message Beep Type</param>
    ''' <returns>Boolean</returns>
    <Runtime.InteropServices.DllImport("USER32.DLL")>
    Public Function MessageBeep(ByVal type As messageBeepType) As Boolean
    End Function

    ''' <summary>
    ''' Make a sound using the 'PC Beep' method, allowing you to specify the frequency and duration of the sound.
    ''' </summary>
    ''' <param name="frequency">Frequency of the Beep</param>
    ''' <param name="duration">Duration of the Beep</param>
    ''' <returns>Boolean</returns>
    <Runtime.InteropServices.DllImport("KERNEL32.DLL")>
    Public Function Beep(ByVal frequency As Integer, ByVal duration As Integer) As Boolean
    End Function
End Module