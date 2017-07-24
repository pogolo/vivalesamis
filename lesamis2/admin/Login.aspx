<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="Login.aspx.vb" Inherits="admin_Login"  %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
	<fieldset><Legend>Login</Legend>
			<table>
					<tr>
						<td width="94">Username</td>
						<td>
							<asp:TextBox id="Username" runat="server" Width="250px"></asp:TextBox></td>
					</tr>
					<tr>
						<td width="94">Password</td>
						<td>
							<asp:TextBox id="Password" runat="server" Width="250px" TextMode="Password"></asp:TextBox></td>
					</tr>
				</table>
<asp:Button id="cmdLogin" runat="server" Text="Login"></asp:Button>&nbsp;
<asp:Button id="cmdCancel" runat="server" Text="Cancel"></asp:Button>
			</fieldset>
</asp:Content>

