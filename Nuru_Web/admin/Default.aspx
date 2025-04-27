<%@ Page Title="" Language="C#" MasterPageFile="~/Styles/NuruWeb.Master" AutoEventWireup="true" CodeBehind="CaptureOrder.aspx.cs" Inherits="Nuru_Web.Pages_Administration.CapturePurchaseOrder" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">


	
<%--	
	<asp:LoginView ID="LoginStatus1" runat="server">
		<AnonymousTemplate>
			<a href="/Simple/login.aspx">Login</a>
		</AnonymousTemplate>
		<LoggedInTemplate>
			<asp:LoginName ID="LoginName1" runat="server" FormatString="Welcome, {0}" />&nbsp;&nbsp;
			<a href="/Simple/logout.aspx">Logout</a>
		</LoggedInTemplate>
	</asp:LoginView>
	<br />
	<br />--%>
	
	<table class="gnav">
	<tr>
		<td><a href="/admin/access/access_rules.aspx">Admin</a></td>
		<td><a href="/admin/management.aspx">Management</a></td>
	</tr>
	</table>

</asp:Content>

