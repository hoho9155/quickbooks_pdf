<%@ Page Language="vb" Async="true" AutoEventWireup="false" CodeBehind="default.aspx.vb" Inherits="PDf_Upload._default" %>

<%@ Register assembly="DevExpress.Web.v22.2, Version=22.2.5.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web" tagprefix="dx" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <dx:ASPxFormLayout ID="frm1" runat="server" ClientInstanceName="frm1" Width="100%">
                <Items>
                    <dx:LayoutGroup Caption="Invoice Importer" ColCount="2" ColSpan="1" ColumnCount="2">
                        <GroupBoxStyle>
                            <Caption Font-Size="16pt" ForeColor="#0066FF">
                            </Caption>
                        </GroupBoxStyle>
                        <Items>
                            <dx:EmptyLayoutItem ColSpan="2" ColumnSpan="2">
                            </dx:EmptyLayoutItem>
                            <dx:LayoutItem Caption="Select and upload your PDF Invoice" ColSpan="1" HorizontalAlign="Left" VerticalAlign="Top" Width="30%">
                                <LayoutItemNestedControlCollection>
                                    <dx:LayoutItemNestedControlContainer runat="server">
                                        <dx:ASPxUploadControl ID="upload1" runat="server" AutoStartUpload="True" ClientInstanceName="upload1" NullText="Browse or drag and drop pdf invoice here." UploadMode="Advanced" Width="100%">
                                            <ValidationSettings AllowedFileExtensions=".pdf">
                                            </ValidationSettings>
                                            <ClientSideEvents FilesUploadStart="function(s, e) {
	loading1.Show();
}" FileUploadComplete="function(s, e) {
loading1.Hide();
	
}" />
                                            <AdvancedModeSettings EnableDragAndDrop="True">
                                            </AdvancedModeSettings>
                                        </dx:ASPxUploadControl>
                                    </dx:LayoutItemNestedControlContainer>
                                </LayoutItemNestedControlCollection>
                                <CaptionSettings Location="Top" />
                                <CaptionStyle Font-Size="14pt">
                                </CaptionStyle>
                            </dx:LayoutItem>
                            <dx:EmptyLayoutItem ColSpan="1" HorizontalAlign="Left" VerticalAlign="Top" Width="70%">
                            </dx:EmptyLayoutItem>
                        </Items>
                    </dx:LayoutGroup>
                </Items>
            </dx:ASPxFormLayout>
        </div>
        <dx:ASPxLoadingPanel ID="loading1" runat="server" ClientInstanceName="loading1" Text="Working , please wait&amp;hellip;">
        </dx:ASPxLoadingPanel>
    </form>
</body>
</html>

