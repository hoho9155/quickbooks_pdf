Imports DevExpress.Web
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Newtonsoft.Json
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Web.UI
Imports System.Configuration
Imports System.Web
Imports Intuit.Ipp.OAuth2PlatformClient
Imports System.Threading.Tasks
Imports Intuit.Ipp.Core
Imports Intuit.Ipp.Data
Imports Intuit.Ipp.DataService
Imports Intuit.Ipp.QueryFilter
Imports Intuit.Ipp.Security
Imports Intuit.Ipp.Exception
Imports System.Linq
Imports Intuit.Ipp.ReportService

Public Class _default
    Inherits System.Web.UI.Page
    Private Const UploadDirectory As String = "~/UploadedFiles/"

#Region "QBO"

    ' OAuth2 client configuration
    Private Shared redirectURI As String = ConfigurationManager.AppSettings("redirectURI")
    Private Shared _clientID As String = ConfigurationManager.AppSettings("clientID")
    Private Shared clientSecret As String = ConfigurationManager.AppSettings("clientSecret")
    Private Shared appEnvironment As String = ConfigurationManager.AppSettings("appEnvironment")
    Private Shared oauthClient
    Private Shared authCode As String
    Private Shared idToken As String
    Public Shared keys As IList(Of JsonWebKey)
    Public Shared dictionary As New Dictionary(Of String, String)()

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Return
        oauthClient = New OAuth2Client(_clientID, clientSecret, redirectURI, appEnvironment)
        Dim tokenObtained = False
        If Not dictionary.ContainsKey("accessToken") Then
            If Request.QueryString.Count > 0 Then
                Dim response = New AuthorizeResponse(Request.QueryString.ToString())
                If response.State IsNot Nothing Then
                    If oauthClient.CSRFToken = response.State Then
                        If response.RealmId IsNot Nothing Then
                            If Not dictionary.ContainsKey("realmId") Then
                                dictionary.Add("realmId", response.RealmId)
                            End If
                        End If

                        If response.Code IsNot Nothing Then
                            authCode = response.Code
                            Dim t As PageAsyncTask = New PageAsyncTask(AddressOf PerformCodeExchange)
                            Page.RegisterAsyncTask(t)
                            Page.ExecuteRegisteredAsyncTasks()
                        End If
                        tokenObtained = True

                    Else
                        Output("Invalid State")
                        dictionary.Clear()
                    End If
                End If
            End If
        End If

        Try
            If Not tokenObtained And Not dictionary.ContainsKey("accessToken") Then
                Dim scopes = New List(Of OidcScopes) From {OidcScopes.Accounting}
                Dim authorizationRequest = oauthClient.GetAuthorizationURL(scopes)
                Response.Redirect(authorizationRequest)
            End If
        Catch ex As Exception
            Output(ex.Message)
        End Try

    End Sub

    ' Start code exchange to get the Access Token and Refresh Token
    Public Async Function PerformCodeExchange() As System.Threading.Tasks.Task
        Try
            Dim tokenResp = Await oauthClient.GetBearerTokenAsync(authCode)
            If Not dictionary.ContainsKey("accessToken") Then
                dictionary.Add("accessToken", tokenResp.AccessToken)
            Else
                dictionary("accessToken") = tokenResp.AccessToken
            End If

            If Not dictionary.ContainsKey("refreshToken") Then
                dictionary.Add("refreshToken", tokenResp.RefreshToken)
            Else
                dictionary("refreshToken") = tokenResp.RefreshToken
            End If

            If tokenResp.IdentityToken IsNot Nothing Then
                idToken = tokenResp.IdentityToken
            End If
            If (Request.Url.Query = "") Then
                Response.Redirect(Request.RawUrl)
            Else
                Response.Redirect(Request.RawUrl.Replace(Request.Url.Query, ""))
            End If
        Catch ex As Exception
            Output("Problem while getting bearer tokens.")
        End Try
    End Function

#End Region


    Public Sub Output(msg As String)

    End Sub

#Region "Upload File"
    Protected Sub Upload1_FileUploadComplete(sender As Object, e As DevExpress.Web.FileUploadCompleteEventArgs) Handles Upload1.FileUploadComplete

        Try
            e.CallbackData = SavePostedFile(e.UploadedFile)
            StartExtract(Session("FileNameAndPath"))
        Catch ex As Exception
            e.IsValid = False
            e.ErrorText = ex.Message

        End Try
    End Sub
    Protected Function SavePostedFile(ByVal uploadedFile As UploadedFile) As String
        Dim ret As String = ""
        Dim FileIdentifier As String = ""
        If uploadedFile.IsValid Then

            Dim fileInfo As New FileInfo(uploadedFile.FileName)
            Session("FileName") = uploadedFile.FileName.ToString
            Dim Seperator = InStr(StrReverse(uploadedFile.FileName), ".")
            Dim fileExtension = StrReverse(Left(StrReverse(Session("FileName")), Seperator))
            Dim MyFileName = Session("FileName")
            Dim caseFilename = Session("FileName")
            Dim resFileName As String = MapPath(UploadDirectory) + FileIdentifier + caseFilename
            uploadedFile.SaveAs(resFileName)
            Session("FileNameAndPath") = resFileName
        Else
        End If
        Return Nothing
    End Function

    Sub StartExtract(ByVal Filename)
        ' Dim FileName As String = "INVOICE-10990.PDF"
        Dim FileBytes = File.ReadAllBytes(Filename)

        Dim cbParser As New CarlisleBrassInvoiceParser()
        Dim mmParser As New MMarcusInvoiceParser()

        If cbParser.ExtractFromPDF(FileBytes) Then
            dictionary.TryGetValue("accessToken", cbParser.AccessToken)
            dictionary.TryGetValue("realmId", cbParser.RealmId)
            Inputs.CreateBill(cbParser)
            ' it means successfull parsing
            ''Process Data, like storing into DB
        ElseIf mmParser.ExtractFromPDF(FileBytes) Then
            dictionary.TryGetValue("accessToken", mmParser.AccessToken)
            dictionary.TryGetValue("realmId", mmParser.RealmId)
            Inputs.CreateBill(mmParser)
            ' it means successfull parsing
            ''Process Data, like storing into DB
        End If
    End Sub
#End Region
End Class