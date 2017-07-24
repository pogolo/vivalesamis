<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="EmailAdmin.aspx.vb" Inherits="admin_EmailAdmin"  %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<br>
			<br>
			<asp:datagrid id="EmailDataGrid" runat="server" Width="100%" AutoGenerateColumns="False">
				<HeaderStyle Font-Bold="True"></HeaderStyle>
				<Columns>
					<asp:BoundColumn DataField="EmailAddress" HeaderText="Email"></asp:BoundColumn>
					<asp:BoundColumn DataField="FullName" HeaderText="Name"></asp:BoundColumn>
					<asp:BoundColumn DataField="Comments" HeaderText="Comments"></asp:BoundColumn>
					<asp:BoundColumn DataField="DateAdded" HeaderText="Date Added"></asp:BoundColumn>
				</Columns>
			</asp:datagrid>&nbsp;&nbsp;
			<asp:Panel id="ClearPanel" runat="server">
<asp:Button id="cmdClear" runat="server" Text="Clear"></asp:Button>&nbsp; 
<asp:Button id="cmdCancel" runat="server" Text="Cancel"></asp:Button></asp:Panel>
			<asp:Panel id="PromptPanel" runat="server" Visible="False">l 
<asp:Label id="Label1" runat="server" BackColor="Transparent" ForeColor="Red">This will delete all emails. Are you sure you want to continue?</asp:Label>
<asp:Button id="cmdYes" runat="server" Text="Yes"></asp:Button>&nbsp; 
<asp:Button id="cmdNo" runat="server" Text="No"></asp:Button></asp:Panel>
</asp:Content>

