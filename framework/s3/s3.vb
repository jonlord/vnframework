Imports Amazon.S3
Imports Amazon.S3.Model
Imports Amazon.S3.Transfer
Imports System.Security

''' <summary>
''' Wrapper class to remove abstraction of S3 functions
''' </summary>
Public Class s3Library
    Private Const ERRORMSG_CREATEINGBUCKET As String = "Error creating bucket: "
    Private Const ERRORMSG_CREATEINGFOLDER As String = "Error creating folder: "
    Private Const ERRORMSG_CHECKFILEEXISTS As String = "Error checking if file exists: "

    Private complete As Boolean = False
    Private client As AmazonS3
    Private config As New TransferUtilityConfig
    Private fileTransferUtility As TransferUtility

    ''' <summary>
    ''' Creates a new instance of the s3 wrapper class
    ''' </summary>
    ''' <param name="accessKey">AWS Access Key</param>
    ''' <param name="secretKey">AWS Secret Key</param>
    Public Sub New(accessKey As SecureString, secretKey As SecureString)
        client = New AmazonS3Client(accessKey.ToString, secretKey.ToString)
    End Sub

    ''' <summary>
    ''' Creates a bucket, continues if it already exists
    ''' </summary>
    ''' <param name="bucketName">Bucket to create</param>
    Public Sub createBucket(ByVal bucketName As String)
        Try
            Dim response As ListBucketsResponse = client.ListBuckets()
            Dim found As Boolean = False
            For Each bucket As S3Bucket In response.Buckets
                If bucket.BucketName = bucketName Then
                    found = True
                    Exit For
                End If
            Next
            If Not found Then client.PutBucket(New PutBucketRequest().WithBucketName(bucketName))
        Catch ex As Exception
            Throw New Exception(ERRORMSG_CREATEINGBUCKET & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Creates a new folder in a specified bucket
    ''' </summary>
    ''' <param name="bucketname">Bucket to create folder in</param>
    ''' <param name="folderName">Folder to create</param>
    Public Sub createNewFolder(ByVal bucketname As String, ByVal folderName As String)
        Try
            Dim request As New PutObjectRequest()
            request.WithBucketName(bucketname)
            request.WithKey(folderName)
            request.WithContentBody("")
            client.PutObject(request)
        Catch ex As Net.WebException
            Throw New Exception(String.Format(ERRORMSG_CREATEINGFOLDER, folderName, ex.Message, bucketname))
        End Try
    End Sub
    ''' <summary>
    ''' Uploads a file
    ''' </summary>
    ''' <param name="bucketName">Bucket to use</param>
    ''' <param name="sourceFileName">Local file to upload (full path)</param>
    ''' <param name="destinationFilename">Remote path to store in (full path)</param>
    ''' <returns>True if upload successful, false if it fails</returns>
    Public Function uploadFile(ByVal bucketName As String, ByVal sourceFileName As String, ByVal destinationFilename As String) As Boolean
        Try
            config.NumberOfUploadThreads = 2
            config.DefaultTimeout = 100000
            fileTransferUtility = New TransferUtility(client, config)
            Dim uploadRequest As New TransferUtilityUploadRequest() With {.Timeout = 3600000, .PartSize = 1024 * 1024 * 6}
            uploadRequest.WithBucketName(bucketName)
            uploadRequest.WithKey(destinationFilename)
            uploadRequest.WithFilePath(sourceFileName)
            fileTransferUtility.Upload(uploadRequest)
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Checks if a file exists
    ''' </summary>
    ''' <param name="bucketName">Bucket name to search in</param>
    ''' <param name="sourceFileName">Local file to upload (full path)</param>
    ''' <returns>True if found, false if not found</returns>
    Public Function fileExists(ByVal bucketName As String, ByVal sourceFileName As String) As Boolean
        Try
            Dim request As New ListObjectsRequest()
            request.WithBucketName(bucketName)
            request.WithPrefix(sourceFileName)
            Dim response As ListObjectsResponse = client.ListObjects(request)
            For Each item As S3Object In response.S3Objects
                If item.Key = sourceFileName Then Return True
            Next
            Return False
        Catch ex As Exception
            Throw New Exception(ERRORMSG_CHECKFILEEXISTS & ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Gets the response callback
    ''' </summary>
    ''' <returns>Sets the complete variable to true</returns>
    Private Function beginGetResponseCallback() As AsyncCallback
        beginGetResponseCallback = Nothing
        complete = True
    End Function
End Class