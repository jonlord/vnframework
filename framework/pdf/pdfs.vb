Imports System.IO
Imports iTextSharp.text.pdf
Imports iTextSharp.text

''' <summary>
''' Wrapper functions that extend iTextSharp PDF Library
''' </summary>
Public Class pdfs
    ''' <summary>
    ''' Gets a PDFs content as text
    ''' </summary>
    ''' <param name="path">Path to PDF file</param>
    ''' <returns>The PDF's content as text</returns>
    Shared Function toText(path As String) As String
        Dim sOut As String = ""
        Dim oReader As New PdfReader(path)

        For i = 1 To oReader.NumberOfPages
            Dim its As New parser.SimpleTextExtractionStrategy
            sOut &= parser.PdfTextExtractor.GetTextFromPage(oReader, i, its)
        Next
        Return sOut
    End Function
    ''' <summary>
    ''' Extracts one page from a PDF File
    ''' </summary>
    ''' <param name="path">Path to PDF file</param>
    ''' <param name="pageNumber">Page number to extract</param>
    ''' <returns>One page of a PDF document as the path to a PDF</returns>
    Shared Function extractPage(path As String, pageNumber As Integer) As String
        Try
            Dim reader As New PdfReader(path)
            Dim tempFileName As String = System.IO.Path.GetTempFileName()

            Dim doc As New Document()
            Dim copy As New PdfCopy(doc, New FileStream(tempFileName, FileMode.OpenOrCreate, FileAccess.Write))
            doc.Open()
            copy.Open()
            Dim page As PdfImportedPage
            If pageNumber <= reader.NumberOfPages Then
                page = copy.GetImportedPage(reader, pageNumber)
                copy.AddPage(page)
            Else
                Throw New ArgumentNullException("pageBytes", "No se ha especificado ningún archivo.")
            End If
            doc.Close()
            copy.Close()
            Return tempFileName
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' Appends one pdf document to another
    ''' </summary>
    ''' <param name="sInFilePath">PDF file to append as a byte array</param>
    ''' <param name="pdfDocument">Path to PDF file</param>
    ''' <param name="oPdfWriter"></param>
    Shared Sub addPdf(ByVal sInFilePath() As Byte, ByRef pdfDocument As Document, ByVal oPdfWriter As PdfWriter)
        Dim oDirectContent As PdfContentByte = oPdfWriter.DirectContent
        Dim oPdfReader As PdfReader = New PdfReader(sInFilePath)
        Dim iNumberOfPages As Integer = oPdfReader.NumberOfPages
        Dim iPage As Integer = 0
        Do While (iPage < iNumberOfPages)
            iPage += 1
            pdfDocument.SetPageSize(oPdfReader.GetPageSizeWithRotation(iPage))
            pdfDocument.NewPage()
            Dim oPdfImportedPage As PdfImportedPage = oPdfWriter.GetImportedPage(oPdfReader, iPage)
            Dim iRotation As Integer = oPdfReader.GetPageRotation(iPage)
            If (iRotation = 90) Or (iRotation = 270) Then
                oDirectContent.AddTemplate(oPdfImportedPage, 0, -1.0F, 1.0F, 0, 0, oPdfReader.GetPageSizeWithRotation(iPage).Height)
            Else
                oDirectContent.AddTemplate(oPdfImportedPage, 1.0F, 0, 0, 1.0F, 0, 0)
            End If
        Loop
        pdfDocument.NewPage()
    End Sub
    ''' <summary>
    ''' Appends an image to a PDF file as a new page
    ''' </summary>
    ''' <param name="sInFilePath">PDF file to append as a byte array</param>
    ''' <param name="pdfdocument">Path to PDF file</param>
    ''' <param name="oPdfWriter"></param>
    Shared Sub addImage(ByVal sInFilePath() As Byte, ByRef pdfDocument As Document, fileExtension As String, ByVal oPdfWriter As PdfWriter)
        Dim rect As Rectangle = Nothing
        Dim X, Y As Single

        Try
            If (sInFilePath Is Nothing) Then _
               Throw New FileNotFoundException

            Dim data As Integer = sInFilePath.Length
            Dim tempFileName As String = Path.GetTempFileName()

            Using fs As New FileStream(tempFileName, FileMode.OpenOrCreate)
                Dim bw As New BinaryWriter(fs)
                bw.Write(sInFilePath, 0, data)
                bw.Flush()
                bw = Nothing
            End Using

            If fileExtension <> "pdf" Then
                Dim img As Image = Image.GetInstance(tempFileName) 'Image path reference
                rect = pdfDocument.PageSize
                If img.Width > rect.Width OrElse img.Height > rect.Height Then
                    img.ScaleToFit(rect.Width, rect.Height)
                    X = (rect.Width - img.ScaledWidth) / 2
                    Y = (rect.Height - img.ScaledHeight) / 2
                Else
                    X = (rect.Width - img.Width) / 2
                    Y = (rect.Height - img.Height) / 2
                End If
                img.SetAbsolutePosition(X, Y)
                pdfDocument.Add(img) 'Add image to document
                pdfDocument.NewPage() 'Add page break
            Else
                addPdf(sInFilePath, pdfDocument, oPdfWriter)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub
End Class