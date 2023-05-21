Imports System.IO
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Canvas.Parser

Public Class CarlisleBrassInvoiceParser
    Inherits InvoiceParser

    ''' <summary>
    ''' Reti
    ''' </summary>
    ''' <param name="Bytes"></param>
    ''' <returns></returns>
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

                Dim InvoiceNumberCaption = ExtractTextByRect(420, 751, 515, 735, strategy1.ResultCoordinates)
                Dim InvoiceDateCaption = ExtractTextByRect(420, 735, 515, 718, strategy1.ResultCoordinates)
                Dim AccountCaption = ExtractTextByRect(420, 718, 515, 701, strategy1.ResultCoordinates)
                Dim OrderReferenceCaption = ExtractTextByRect(420, 701, 515, 683, strategy1.ResultCoordinates)

                Dim ProductCodeCaption = ExtractTextByRect(31, 528, 122, 515, strategy1.ResultCoordinates)
                Dim DescriptionOfGoodsCaption = ExtractTextByRect(122, 528, 390, 515, strategy1.ResultCoordinates)
                Dim QuantityCaption = ExtractTextByRect(390, 528, 448, 515, strategy1.ResultCoordinates)
                Dim PricePerItemCaption = ExtractTextByRect(448, 528, 546, 515, strategy1.ResultCoordinates)
                Dim TotalRowCaption = ExtractTextByRect(546, 528, 579, 515, strategy1.ResultCoordinates)

                Dim Total2Caption = ExtractTextByRect(430, 88, 497, 75, strategy1.ResultCoordinates)
                Dim VatCaption = ExtractTextByRect(430, 75, 497, 63, strategy1.ResultCoordinates)
                Dim InvoiceTotalCaption = ExtractTextByRect(430, 63, 497, 51, strategy1.ResultCoordinates)


                If InvoiceNumberCaption = "INVOICE NUMBER" AndAlso
                    InvoiceDateCaption = "INVOICE DATE" AndAlso
                    AccountCaption = "ACCOUNT" AndAlso
                    OrderReferenceCaption = "ORDER REFERENCE" AndAlso
                    ProductCodeCaption = "PRODUCT CODE" AndAlso
                    DescriptionOfGoodsCaption = "DESCRIPTION OF GOODS" AndAlso
                    QuantityCaption = "QUANTITY" AndAlso
                    PricePerItemCaption = "PRICE PER ITEM" AndAlso
                    TotalRowCaption = "Total" AndAlso
                    Total2Caption = "TOTAL" AndAlso
                    VatCaption = "VAT at 20.00 %" AndAlso
                    InvoiceTotalCaption = "INVOICE TOTAL" AndAlso
                    strategy1.AccumulatedText.Contains("CARLISLEBRASSLTD") Then



                    InvoiceNumber = ExtractTextByRect(516, 751, 578, 735, strategy1.ResultCoordinates)
                    InvoiceDate = ExtractTextByRect(516, 735, 578, 718, strategy1.ResultCoordinates)
                    Account = ExtractTextByRect(516, 718, 578, 701, strategy1.ResultCoordinates)
                    OrderReference = ExtractTextByRect(516, 701, 578, 683, strategy1.ResultCoordinates)

                    Dim StartY = 515
                    Dim Items As New List(Of InvoiceProduct)

                    For i = 1 To 35 ' Consider 35 rows maximum
                        Dim ProductCode = ExtractTextByRect(31, StartY, 122, StartY - 12, strategy1.ResultCoordinates).Trim()
                        Dim DescriptionOfGoods = ExtractTextByRect(122, StartY, 390, StartY - 12, strategy1.ResultCoordinates).Trim()
                        Dim Quantity = ExtractTextByRect(390, StartY, 448, StartY - 12, strategy1.ResultCoordinates).Trim()
                        Dim PricePerItem = ExtractTextByRect(448, StartY, 546, StartY - 12, strategy1.ResultCoordinates).Trim()
                        Dim TotalRow = ExtractTextByRect(546, StartY, 579, StartY - 12, strategy1.ResultCoordinates).Trim()

                        If Not String.IsNullOrEmpty(ProductCode) OrElse Not String.IsNullOrEmpty(DescriptionOfGoods) OrElse Not String.IsNullOrEmpty(Quantity) OrElse Not String.IsNullOrEmpty(PricePerItem) OrElse Not String.IsNullOrEmpty(TotalRow) Then
                            Dim NewProduct As New InvoiceProduct
                            NewProduct.ProductCode = ProductCode
                            NewProduct.Description = DescriptionOfGoods
                            NewProduct.Quantity = Convert.ToDouble(Quantity)
                            NewProduct.PricePerItem = Convert.ToDouble(PricePerItem.Replace("£", ""))
                            NewProduct.TotalRow = Convert.ToDouble(TotalRow.Replace("£", ""))
                            Items.Add(NewProduct)
                        End If
                        StartY = StartY - 12
                    Next

                    Products = Items.ToArray()

                    Total = Convert.ToDouble(ExtractTextByRect(497, 88, 579, 75, strategy1.ResultCoordinates).Replace("£", ""))
                    VAT = Convert.ToDouble(ExtractTextByRect(497, 75, 579, 63, strategy1.ResultCoordinates).Replace("£", ""))
                    InvoiceTotal = Convert.ToDouble(ExtractTextByRect(497, 63, 579, 51, strategy1.ResultCoordinates).Replace("£", ""))
                    Result = True
                End If
            End Using
        Catch ex As Exception

        End Try
        Return Result
    End Function
End Class
