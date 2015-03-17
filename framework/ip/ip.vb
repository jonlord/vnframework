Imports System.Net

''' <summary>
''' IP address manipulation library
''' </summary>
Public Class ip
    ''' <summary>
    ''' Supported protocol versions
    ''' </summary>
    Enum protocolVersion
        IPv4 = 1
        IPv6 = 2
    End Enum
    ''' <summary>
    ''' Gets the local computer's IP Address
    ''' </summary>
    ''' <param name="desiredSubnet">If multiple IP address are present, the function gives preference to a particular subnet</param>
    ''' <param name="protocolVersion">IPv4 or IPv6 Protocol</param>
    ''' <returns>Matching IP address</returns>
    Shared Function GetIPAddress(desiredSubnet As String, protocolVersion As protocolVersion) As String
        GetIPAddress = ""
        Dim strHostName As String = Dns.GetHostName()
        For Each ipAddress As IPAddress In Dns.GetHostEntry(strHostName).AddressList
            Dim strIPAddress As String = ipAddress.ToString()
            Dim isDesiredProtocol As Boolean = IIf((strIPAddress.IndexOf(".") > 0 And protocolVersion = protocolVersion.IPv4) Or (strIPAddress.IndexOf(":") > 0 And protocolVersion = protocolVersion.IPv6), True, False)
            If isDesiredProtocol AndAlso (GetIPAddress = "" Or strIPAddress.StartsWith(desiredSubnet)) Then
                GetIPAddress = strIPAddress
                Exit For
            End If
        Next
    End Function
    ''' <summary>
    ''' Gets the lastOctet of an IP Addess
    ''' </summary>
    ''' <param name="desiredSubnet">If multiple IP address are present, the function gives preference to a particular subnet</param>
    ''' <param name="protocolVersion">IPv4 or IPv6 Protocol</param>
    ''' <returns>Last octet as an integer</returns>
    Shared Function lastOctetIPAddress(desiredSubnet As String, protocolVersion As protocolVersion) As Integer
        Dim lastOctet As String = GetIPAddress(desiredSubnet, protocolVersion)
        Dim parts() As String = lastOctet.Split(".")
        lastOctetIPAddress = parts(UBound(parts) - 1)
    End Function
End Class