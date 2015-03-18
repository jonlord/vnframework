''' <summary>
''' Opens a simple form to navigate through images
''' </summary>
Public Class frmImageBrowser
    ''' <summary>
    ''' ArrayList containing images
    ''' </summary>
    Private imageArray As ArrayList
    ''' <summary>
    ''' Zero Based position in the arraylist
    ''' </summary>
    Private pos As Integer = 0

    ''' <summary>
    ''' Loads the form with the first image
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub frmImageBrowser_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If imageArray.Count > pos Then pbBrowser.Image = imageArray(pos)
    End Sub

    ''' <summary>
    ''' Initializes the form with a list of images stored in an ArrayList
    ''' </summary>
    ''' <param name="imageArray">ArrayList containing images</param>
    Public Sub New(imageArray As ArrayList)
        InitializeComponent()
        Me.imageArray = imageArray
    End Sub

    ''' <summary>
    ''' Gets Previous image in ArrayList
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmdPrev_Click(sender As Object, e As EventArgs) Handles cmdPrev.Click
        If pos > 0 Then pos -= 1
        If imageArray.Count > pos Then pbBrowser.Image = imageArray(pos)
    End Sub

    ''' <summary>
    ''' Get Next Image in ArrayList
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmdNext_Click(sender As Object, e As EventArgs) Handles cmdNext.Click
        If pos + 1 < imageArray.Count Then pos += 1
        If imageArray.Count > pos Then pbBrowser.Image = imageArray(pos)
    End Sub
End Class