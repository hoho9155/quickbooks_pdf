Imports System.IO
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Canvas.Parser

Public Class MMarcusInvoiceParser
    Inherits InvoiceParser

    Public Overrides Function ExtractFromPDF(Bytes() As Byte) As Boolean
        Dim Result As Boolean = False
        Try
            Using ms As New MemoryStream(Bytes)
                Dim reader As New PdfReader(ms)

                Dim document As New PdfDocument(reader)

                Dim strategy1 As New PdfTextLocator()
                Dim raw_text As String

                For i = 1 To document.GetNumberOfPages()
                    Dim page = document.GetPage(i)

                    raw_text = PdfTextExtractor.GetTextFromPage(page, strategy1)

                    Dim pageSize = page.GetPageSize()
                Next

                Dim InvoiceNumberCaption = ExtractTextByRect(530, 614, 578, 587, strategy1.ResultCoordinates)
                Dim InvoiceDateCaption = ExtractTextByRect(14, 614, 61, 587, strategy1.ResultCoordinates)
                Dim AccountCaption = ExtractTextByRect(468, 614, 525, 587, strategy1.ResultCoordinates)
                Dim OrderReferenceCaption = ExtractTextByRect(366, 614, 464, 587, strategy1.ResultCoordinates)

                Dim ProductCodeCaption = ExtractTextByRect(38, 555, 130, 533, strategy1.ResultCoordinates)
                Dim DescriptionOfGoodsCaption = ExtractTextByRect(150, 555, 300, 533, strategy1.ResultCoordinates)
                Dim QuantityCaption = ExtractTextByRect(320, 555, 355, 533, strategy1.ResultCoordinates)
                Dim PricePerItemCaption = ExtractTextByRect(470, 555, 510, 533, strategy1.ResultCoordinates)
                Dim TotalRowCaption = ExtractTextByRect(530, 555, 578, 533, strategy1.ResultCoordinates)

                Dim Total2Caption = ExtractTextByRect(13, 86, 80, 62, strategy1.ResultCoordinates)
                Dim VatCaption = ExtractTextByRect(150, 86, 190, 62, strategy1.ResultCoordinates)
                Dim InvoiceTotalCaption = ExtractTextByRect(320, 86, 470, 62, strategy1.ResultCoordinates)

                If InvoiceNumberCaption = "INVOICENo." AndAlso
                    InvoiceDateCaption = "INVOICEDATE" AndAlso
                    AccountCaption = "ACCOUNTNo." AndAlso
                    OrderReferenceCaption = "YOURORDER No." AndAlso
                    ProductCodeCaption = "PRODUCT No." AndAlso
                    DescriptionOfGoodsCaption = "DESCRIPTION" AndAlso
                    QuantityCaption = "QTY" AndAlso
                    PricePerItemCaption = "NETT" AndAlso
                    TotalRowCaption = "GROSS" AndAlso
                    Total2Caption = "TOTAL GOODS" AndAlso
                    VatCaption = "VAT @%20.00" AndAlso
                    InvoiceTotalCaption = "INVOICE TOTAL" AndAlso
                    strategy1.AccumulatedText.Contains("m. marcus") Then

                    InvoiceNumber = ExtractTextByRect(530, 580, 578, 554, strategy1.ResultCoordinates)
                    InvoiceDate = ExtractTextByRect(14, 580, 61, 554, strategy1.ResultCoordinates)
                    Account = ExtractTextByRect(468, 580, 525, 554, strategy1.ResultCoordinates)
                    OrderReference = ExtractTextByRect(366, 580, 464, 554, strategy1.ResultCoordinates)


                    Dim StartY = 518
                    Dim Items As New List(Of InvoiceProduct)

                    For i = 1 To 35 ' Consider 35 rows maximum
                        Dim ProductCode = ExtractTextByRect(50, StartY, 140, StartY - 12, strategy1.ResultCoordinates).Trim()
                        Dim DescriptionOfGoods = ExtractTextByRect(150, StartY, 305, StartY - 12, strategy1.ResultCoordinates).Trim()
                        Dim Quantity = ExtractTextByRect(327, StartY, 355, StartY - 12, strategy1.ResultCoordinates).Trim()
                        Dim PricePerItem = ExtractTextByRect(480, StartY, 546, StartY - 12, strategy1.ResultCoordinates).Trim()
                        Dim TotalRow = ExtractTextByRect(550, StartY, 579, StartY - 12, strategy1.ResultCoordinates).Trim()

                        If Not String.IsNullOrEmpty(ProductCode) OrElse Not String.IsNullOrEmpty(DescriptionOfGoods) OrElse Not String.IsNullOrEmpty(Quantity) OrElse Not String.IsNullOrEmpty(PricePerItem) OrElse Not String.IsNullOrEmpty(TotalRow) Then
                            Dim NewProduct As New InvoiceProduct
                            NewProduct.ProductCode = ProductCode
                            NewProduct.Description = DescriptionOfGoods
                            NewProduct.Quantity = Convert.ToDouble(Quantity)
                            NewProduct.PricePerItem = Convert.ToDouble(PricePerItem)
                            NewProduct.TotalRow = Convert.ToDouble(TotalRow)
                            Items.Add(NewProduct)
                        End If
                        StartY = StartY - 12
                    Next

                    Products = Items.ToArray()

                    Total = Convert.ToDouble(ExtractTextByRect(90, 86, 140, 62, strategy1.ResultCoordinates))
                    VAT = Convert.ToDouble(ExtractTextByRect(200, 73, 310, 62, strategy1.ResultCoordinates))
                    InvoiceTotal = Convert.ToDouble(ExtractTextByRect(475, 86, 579, 62, strategy1.ResultCoordinates))

                    Result = True
                End If
            End Using
        Catch ex As Exception

        End Try
        Return Result
    End Function
End Class
