Imports System.Text
Imports iText.Kernel.Geom
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Canvas.Parser
Imports iText.Kernel.Pdf.Canvas.Parser.Data
Imports iText.Kernel.Pdf.Canvas.Parser.Listener

Public Class InvoiceProduct
    Public Property ProductCode As String
    Public Property Description As String
    Public Property Quantity As Double
    Public Property PricePerItem As Double
    Public Property TotalRow As Double
End Class

Public Class MyTextChunk
    Public Property PageNumber As Integer
    Public Property Text As String
    Public Property ResultCoordinates As Rectangle

    Public Sub New(ByVal s As String, ByVal r As Rectangle)
        Text = s
        ResultCoordinates = r
    End Sub
End Class

Public Class PdfTextLocator
    Inherits LocationTextExtractionStrategy

    Public Property TextToSearchFor As String
    Public Property ResultCoordinates As List(Of MyTextChunk)

    Public Property AccumulatedText As String

    Public Shared Function GetTextCoordinates(ByVal page As PdfPage, ByVal s As String) As Rectangle
        Dim strat As PdfTextLocator = New PdfTextLocator()
        PdfTextExtractor.GetTextFromPage(page, strat)

        For Each c As MyTextChunk In strat.ResultCoordinates
            If c.Text = s Then Return c.ResultCoordinates
        Next

        Return Nothing
    End Function

    Public Sub New()
        ResultCoordinates = New List(Of MyTextChunk)()
    End Sub

    Public Overrides Sub EventOccurred(ByVal data As IEventData, ByVal type As EventType)
        If Not type.Equals(EventType.RENDER_TEXT) Then Return
        Dim renderInfo As TextRenderInfo = CType(data, TextRenderInfo)
        Dim text As IList(Of TextRenderInfo) = renderInfo.GetCharacterRenderInfos()

        For i As Integer = 0 To text.Count - 1

            Dim startX As Single = text(i).GetBaseline().GetStartPoint().[Get](0)
            Dim startY As Single = text(i).GetBaseline().GetStartPoint().[Get](1)

            'If text(i).GetText() = TextToSearchFor(0).ToString() Then
            'Dim word As String = ""
            'Dim j As Integer = i

            'While j < i + TextToSearchFor.Length AndAlso j < text.Count
            'word = word & text(j).GetText()
            'j += 1
            'End While


            ResultCoordinates.Add(New MyTextChunk(text(i).GetText(), New Rectangle(startX, startY, text(i).GetAscentLine().GetEndPoint().[Get](0) - startX, text(i).GetAscentLine().GetEndPoint().[Get](0) - startY)))
            'End If
            'Debug.WriteLine(String.Format("X {0:N5}, Y {1:N5}", startX, startY))
            AccumulatedText = AccumulatedText & text(i).GetText()
            'Console.WriteLine(text(i).GetText() & " | " & String.Format("X {0:N5}, Y {1:N5}", startX, startY))

        Next
    End Sub
End Class



Public MustInherit Class InvoiceParser
    Inherits QBOData

    Public Property InvoiceNumber As String
    Public Property InvoiceDate As String
    Public Property Account As String
    Public Property OrderReference As String

    Public Products() As InvoiceProduct

    Public Property Total As Double
    Public Property VAT As Double
    Public Property InvoiceTotal As Double

    Protected Function ExtractTextByRect(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, ByVal Coordinates As List(Of MyTextChunk)) As String
        Dim InsideRect As New List(Of MyTextChunk)
        For Each Chunk In Coordinates
            If Chunk.ResultCoordinates.GetX() >= Left AndAlso Chunk.ResultCoordinates.GetX() <= Right AndAlso Chunk.ResultCoordinates.GetY() <= Top AndAlso Chunk.ResultCoordinates.GetY() >= Bottom Then
                InsideRect.Add(Chunk)
            End If
        Next

        ' Sort if needed

        Dim sb As New StringBuilder
        For Each Chunk In InsideRect
            sb.Append(Chunk.Text)
        Next

        Return sb.ToString().Trim()

    End Function

    Public MustOverride Function ExtractFromPDF(ByVal Bytes() As Byte) As Boolean

End Class
