Imports System.Net
Imports System.IO

''' <summary>
''' Programatically consume a web service using .NET
''' </summary>
Public Class webServiceConsumer
    ''' <summary>
    ''' Parameter for the web service call
    ''' </summary>
    Public Class webServiceParameter
        Public name As String
        Public value As String
        Sub New(name As String, value As String)
            Me.name = name
            Me.value = value
        End Sub
    End Class

    ''' <summary>
    ''' Response type conversion of web service calls
    ''' </summary>
    Public Enum responseType
        stringType = 0
        xmlType = 1
    End Enum

    ''' <summary>
    ''' Builds an XML string with all the parameters and their respective values
    ''' </summary>
    ''' <param name="parameterList">Arraylist containing all the parameters needded for the funcion</param>
    ''' <returns>String containg the XML format</returns>
    Private Shared Function buildParameters(parameterList As ArrayList) As String
        buildParameters = ""
        For Each parameter As webServiceParameter In parameterList
            buildParameters &= String.Format("<{0}>{1}</{0}>", parameter.name, parameter.value)
        Next
    End Function

    ''' <summary>
    ''' Returns the response of a web service call in the desired format
    ''' </summary>
    ''' <param name="url">Web Service URL</param>
    ''' <param name="functionName">Name of the web service function to invoke</param>
    ''' <param name="parameterList">Arraylist containing all the parameters needded for the funcion</param>
    ''' <param name="responseType">Desired response type</param>
    ''' <returns>Web servicr response in the desired response type</returns>
    Public Shared Function getResponse(url As String, functionName As String, parameterList As ArrayList, responseType As responseType)
        Select Case responseType
            Case responseType.stringType
                Return getStringResponse(url, functionName, parameterList)
            Case responseType.xmlType
                Return getXMLResponse(url, functionName, parameterList)
            Case Else
                Return getStringResponse(url, functionName, parameterList)
        End Select
    End Function

    ''' <summary>
    ''' Returns the response of a webservice call as string
    ''' </summary>
    ''' <param name="url">Web Service URL</param>
    ''' <param name="functionName">Name of the web service function to invoke</param>
    ''' <param name="parameterList">Arraylist containing all the parameters needded for the funcion</param>
    ''' <returns>String containing the response</returns>
    Private Shared Function getStringResponse(url As String, functionName As String, parameterList As ArrayList) As String
        Dim xml As Xml.XmlDocument = getXML(url, functionName, parameterList)
        Dim xmlNodes As Xml.XmlNodeList = xml.SelectNodes("/")
        For Each node As Xml.XmlNode In xmlNodes
            Return node.InnerText
        Next
        Return -1
    End Function

    ''' <summary>
    ''' Returns the response of a webservice call as an XML Document
    ''' </summary>
    ''' <param name="url">Web Service URL</param>
    ''' <param name="functionName">Name of the web service function to invoke</param>
    ''' <param name="parameterList">Arraylist containing all the parameters needded for the funcion</param>
    ''' <returns>XML Document</returns>
    Private Shared Function getXMLResponse(url As String, functionName As String, parameterList As ArrayList) As Xml.XmlDocument
        Return getXML(url, functionName, parameterList)
    End Function

    ''' <summary>
    ''' Returns the response of a webservice call as an XML Document
    ''' </summary>
    ''' <param name="url">Web Service URL</param>
    ''' <param name="functionName">Name of the web service function to invoke</param>
    ''' <param name="parameterName">Arraylist containing all the parameters needded for the funcion</param>
    ''' <returns>XML Document</returns>
    Private Shared Function getXML(url As String, functionName As String, parameterName As ArrayList) As Xml.XmlDocument
        Dim soap As String = "<?xml version=""1.0"" encoding=""utf-8""?>" &
         "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " &
          "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " &
          "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &
          "<soap:Body>" &
           "<" & functionName & " xmlns=""http://tempuri.org/"">" &
            buildParameters(parameterName) &
           "</" & functionName & ">" &
          "</soap:Body>" &
         "</soap:Envelope>"

        Dim req As HttpWebRequest = WebRequest.Create(url)
        req.Headers.Add("SOAPAction", """http://tempuri.org/" & functionName & """")
        req.ContentType = "text/xml;charset=""utf-8"""
        req.Accept = "text/xml"
        req.Method = "POST"

        Using stm As Stream = req.GetRequestStream()
            Using stmw As StreamWriter = New StreamWriter(stm)
                stmw.Write(soap)
            End Using
        End Using

        Dim response As WebResponse = req.GetResponse()
        Dim responseStream As New StreamReader(response.GetResponseStream())
        soap = responseStream.ReadToEnd

        Dim xml As New Xml.XmlDocument
        xml.LoadXml(Trim(soap))
        Return xml
    End Function
End Class