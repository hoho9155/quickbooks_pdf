Imports System
Imports System.Linq
Imports Intuit.Ipp.Data
Imports Intuit.Ipp.DataService
Imports Intuit.Ipp.QueryFilter
Imports Intuit.Ipp.Security
Imports Intuit.Ipp.Core
Imports DevExpress.XtraRichEdit.Model.History

Public Class Inputs


    Public Shared Function CreateBill(invoiceParser As InvoiceParser)
        ' Step 1: Initialize OAuth2RequestValidator and ServiceContext

        Dim ouathValidator As New OAuth2RequestValidator(invoiceParser.AccessToken)
        Dim serviceContext As New ServiceContext(invoiceParser.RealmId, IntuitServicesType.QBO, ouathValidator)
        'serviceContext.IppConfiguration.BaseUrl.Qbo = "https://sandbox-quickbooks.api.intuit.com"
        serviceContext.IppConfiguration.BaseUrl.Qbo = "https://quickbooks.api.intuit.com/" 'prod
        serviceContext.IppConfiguration.MinorVersion.Qbo = "65"
        Dim dataService = New DataService(serviceContext)


        'Create a new Bill object
        Dim bill As Bill = New Bill()

        ' Set Vendor
        Try
            Dim querySvcVendor = New QueryService(Of Vendor)(serviceContext)
            Dim vendor = querySvcVendor.ExecuteIdsQuery("select * from vendor where DisplayName='Carlisle Brass Ltd'").FirstOrDefault()
            bill.VendorRef = New ReferenceType() With {
                .name = vendor.DisplayName,
                .Value = vendor.Id
            }
            ' Set TxnDate
            bill.TxnDate = Date.ParseExact(invoiceParser.InvoiceDate, "dd/MM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            bill.TxnDateSpecified = True
        Catch ex As Exception
            Throw New Exception("Set Vendor:" + ex.Message)
        End Try


        ' Set GlobalTaxCalculation
        bill.GlobalTaxCalculation = GlobalTaxCalculationEnum.TaxExcluded
        bill.GlobalTaxCalculationSpecified = True

        ' Set the Bill no        
        bill.DocNumber = invoiceParser.InvoiceNumber

        ' Set Terms
        Try
            Dim querySvcTerm As New QueryService(Of Term)(serviceContext)
            Dim terms = querySvcTerm.ExecuteIdsQuery("SELECT * FROM Term").ToList()
            Dim term = terms.Where(Function(t) t.Name = "Net 30").FirstOrDefault()
            If term Is Nothing Then term = terms.FirstOrDefault()
            bill.SalesTermRef = New ReferenceType() With {
            .name = term.Name,
            .Value = term.Id
        }
        Catch ex As Exception
            Throw New Exception("Set Terms:" + ex.Message)
        End Try

        ' Order refrerence
        Dim linked = New List(Of LinkedTxn)()
        linked.Add(New LinkedTxn() With {
            .TxnId = invoiceParser.OrderReference,
            .TxnType = objectNameEnumType.Invoice.ToString() ' The type of the transaction (e.g. Invoice, SalesReceipt)
        })
        bill.LinkedTxn = linked.ToArray()
        bill.PrivateNote = "Order Reference: " + invoiceParser.OrderReference

        ' Get TaxCode
        Dim code As TaxCode
        Try
            Dim querySvcTaxCode As New QueryService(Of TaxCode)(serviceContext)
            Dim codes = querySvcTaxCode.ExecuteIdsQuery("SELECT * FROM TaxCode").ToList()
            code = codes.Where(Function(c) c.Name = "20.0% S").FirstOrDefault()
            If code Is Nothing Then code = codes.FirstOrDefault()
        Catch ex As Exception
            Throw New Exception("Get TaxCode:" + ex.Message)
        End Try

        ' Get Products
        Dim querySvcItem = New QueryService(Of Item)(serviceContext)
        Dim accounts As List(Of Account)

        Try
            ' Get Accounts
            Dim querySvgAccount = New QueryService(Of Account)(serviceContext)
            accounts = querySvgAccount.ExecuteIdsQuery("SELECT * FROM Account").ToList()
        Catch ex As Exception
            Throw New Exception("Get Products:" + ex.Message)
        End Try


        Dim lines = New List(Of Line)
        For Each p As InvoiceProduct In invoiceParser.Products
            If p.ProductCode = String.Empty OrElse p.ProductCode Is Nothing Then
                Dim account = accounts.Where(Function(a) a.Name = "Courier and delivery charges").FirstOrDefault()
                lines.Add(New Line() With {
                    .Amount = p.PricePerItem * p.Quantity,
                    .AmountSpecified = True,
                    .DetailType = LineDetailTypeEnum.AccountBasedExpenseLineDetail,
                    .DetailTypeSpecified = True,
                    .Description = account.Name,
                    .AnyIntuitObject = New AccountBasedExpenseLineDetail() With {
                        .AccountRef = New ReferenceType() With {
                            .name = account.Name,
                            .Value = account.Id
                        },
                        .TaxCodeRef = New ReferenceType() With {
                            .name = code.Name,
                            .Value = code.Id
                        }
                    }
                })
            Else
                Dim item = querySvcItem.ExecuteIdsQuery("SELECT * FROM Item WHERE Sku='" + p.ProductCode + "'").ToList().FirstOrDefault()
                If Not (item Is Nothing) Then
                    lines.Add(New Line() With {
                        .Amount = p.PricePerItem * p.Quantity,
                        .AmountSpecified = True,
                        .DetailType = LineDetailTypeEnum.ItemBasedExpenseLineDetail,
                        .DetailTypeSpecified = True,
                        .Description = $"{p.ProductCode} {p.Description}",
                        .AnyIntuitObject = New ItemBasedExpenseLineDetail() With {
                            .ItemRef = New ReferenceType() With {
                                .name = item.Name,
                                .Value = item.Id
                            },
                            .Qty = p.Quantity,
                            .QtySpecified = True,
                            .TaxCodeRef = New ReferenceType() With {
                                .name = code.Name,
                                .Value = code.Id
                            }
                        }
                    })
                End If
            End If
        Next
        bill.Line = lines.ToArray()

        Try
            dataService.Add(bill)
        Catch ex As Exception
            Throw New Exception("Add a bill:" + ex.Message)
        End Try

        Return True
    End Function

End Class
