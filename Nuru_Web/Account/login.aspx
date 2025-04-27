<%@ Page Language="C#" MasterPageFile="~/Styles/NuruWeb.Master" %>
<script runat="server">
	private void Page_Load()
	{
		Login1.Focus();
	}


</script>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" Runat="Server">
    <asp:Login ID="Login1" runat="server" BackColor="#D9EBF2" BorderColor="#666666" 
        BorderStyle="Solid" BorderWidth="1px" Font-Names="Verdana" Font-Size="10pt" >
		<TitleTextStyle BackColor="#4444AA" Font-Bold="True" ForeColor="#FFFFFF" />
	    <LayoutTemplate>
            <table border="0" cellpadding="1" cellspacing="0" 
                style="border-collapse:collapse;">
                <tr>
                    <td>
                        <table border="0" cellpadding="0">
                            <tr>
                                <td align="center" colspan="2" 
                                    style="color:White;background-color:#4444AA;font-weight:bold;">
                                    Log In</td>
                            </tr>
                            <tr>
                                <td align="right">
                                    <asp:Label ID="UserNameLabel" runat="server" AssociatedControlID="UserName">User name:</asp:Label>
                                </td>
                                <td>
                                    <asp:TextBox ID="UserName" runat="server"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="UserNameRequired" runat="server" 
                                        ControlToValidate="UserName" ErrorMessage="User Name is required." 
                                        ToolTip="User Name is required." ValidationGroup="Login1">*</asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td align="right">
                                    <asp:Label ID="PasswordLabel" runat="server" AssociatedControlID="Password">Password:</asp:Label>
                                </td>
                                <td>
                                    <asp:TextBox ID="Password" runat="server" TextMode="Password"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="PasswordRequired" runat="server" 
                                        ControlToValidate="Password" ErrorMessage="Password is required." 
                                        ToolTip="Password is required." ValidationGroup="Login1">*</asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <asp:CheckBox ID="RememberMe" runat="server" Text="Remember me next time." />
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="2" style="color:Red;">
                                    <asp:Literal ID="FailureText" runat="server" EnableViewState="False"></asp:Literal>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" colspan="2">
                                    <asp:Button ID="LoginButton" runat="server" CommandName="Login" Text="Log In" 
                                        ValidationGroup="Login1" />
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </LayoutTemplate>
	</asp:Login>
	
<%--	<h2>
	    You can log into the system as an Administrator using the username and password "Dan Clem" and "dan" (omitting
	    the quotation marks).	    
	</h2>
	<h2>
	    You can log into the system as a non-Administrator using any of the following user names:
	    "Edward Eel", "Franklin Forester", "Gordy Gordon", "Harold Houdini", or "Ike Iverson". The
	    password for each of these users is their first name, all lower case.
	</h2>
--%></asp:Content>

