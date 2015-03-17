''' <summary>
''' Aids at making items to be used in comboboxes, listboxes, etc.
''' Stores a display name and a value
''' </summary>
Public Class listItem
    Private mId As String
    Private mName As String
    Private mTag As Object

#Region "Constructor(s)"
    Public Sub New(ByVal id As String, ByVal name As String, Optional ByVal tag As Object = Nothing)
        [mId] = id
        mName = name
        mTag = tag
    End Sub

    Public Sub New()
        'Empty default constructor
    End Sub
#End Region
#Region "Properties"

    Public Property id() As String
        Get
            Return mId
        End Get
        Set(ByVal Value As String)
            [mId] = Value
        End Set
    End Property

    Public Property Name() As String
        Get
            Return mName
        End Get
        Set(ByVal Value As String)
            mName = Value
        End Set
    End Property

    Public Property tag() As Object
        Get
            Return mTag
        End Get
        Set(ByVal Value As Object)
            mTag = Value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return mName
    End Function
#End Region
End Class