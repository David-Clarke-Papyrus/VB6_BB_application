<%@ Page Title="" Language="C#" MasterPageFile="~/Styles/NuruWeb.Master" AutoEventWireup="true" CodeBehind="index.aspx.cs" Inherits="Nuru_Web.Default" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManagerProxy ID="ScriptManagerProxy1" runat="server">
    </asp:ScriptManagerProxy>
    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
    <ContentTemplate>
        <div>
                <asp:Label runat="server" ID= "TESTLABEL" Text="Please log in to use this application">
                </asp:Label>
       </div>
    </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
