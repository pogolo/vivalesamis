<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="Contact.aspx.vb" Inherits="Contact"  %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<table ID="Table1">
				<tr>
					<td colspan="2">
						<h1>Contact Information</h1>
					</td>
				</tr>
				<tr>
					<td width="200"></td>
					<td width="400">
						<asp:Panel id="MainPanel" runat="server">
							<FIELDSET><LEGEND>Email</LEGEND>Please join the e-mail list for Viva les Amis to 
								get news of upcoming events.<BR>
								<BR>
								<TABLE>
									<TR>
										<TD>Name</TD>
										<TD>
											<asp:TextBox id="Name" runat="server" Width="250px"></asp:TextBox>*
											<BR>
										</TD>
									</TR>
									<TR>
										<TD>Email</TD>
										<TD>
											<asp:TextBox id="Email" runat="server" Width="250px"></asp:TextBox>*
										</TD>
									</TR>
									<TR>
										<TD>Comments</TD>
										<TD>
											<asp:TextBox id="Comments" runat="server" TextMode="MultiLine" Columns="50" Rows="5"></asp:TextBox></TD>
									</TR>
									<TR>
										<TD colSpan="2">
											<asp:Button id="cmdSubmit" runat="server" Text="Submit"></asp:Button></TD>
									</TR>
								</TABLE>
								For more information: <A href="mailto:info@vivalesamis.com">info@vivalesamis.com</A>
							</FIELDSET>
						</asp:Panel>
    <asp:Panel ID="pnlDone" runat="server" Width="100%" Visible="False">
        <br />
        <br />
        <br />
        Thank you for submitting your email address!</asp:Panel>
					</td>
				</tr>
			</table>
			<br /><br /><br /><br />
</asp:Content>

