<%@ Control Language="VB" AutoEventWireup="false" CodeFile="MenuControl.ascx.vb" Inherits="Controls_MenuControl" %>
<div class="MenuContent">
<div class="MenuLinks">
<table align="center">
    <tr>
        <td width="65"></td>
        <td><asp:HyperLink ID="hlDVD" runat="server" CssClass="DVDLink" NavigateUrl="~/merchandise.aspx">Buy a DVD!</asp:HyperLink></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
    </tr>
       <tr>
        <td><asp:HyperLink ID="hlHome" runat="server" NavigateUrl="~/PageOne.aspx">Les Amis</asp:HyperLink></td> 
        <td><asp:HyperLink ID="hlMerch" runat="server" NavigateUrl="~/merchandise.aspx">Merchandise</asp:HyperLink></td>
        <td><asp:HyperLink ID="hlSynopsis" runat="server" NavigateUrl="~/synopsis.aspx">Synopsis</asp:HyperLink></td>
        <td><asp:HyperLink ID="hlHistory" runat="server" NavigateUrl="~/history.aspx">History</asp:HyperLink></td>
        <td><asp:HyperLink ID="hlGallery" runat="server" NavigateUrl="~/gallery.aspx">Gallery</asp:HyperLink></td>
        <td><asp:HyperLink ID="hlTrailer" runat="server" NavigateUrl="~/trailer.aspx">Trailer</asp:HyperLink></td>
        <td><asp:HyperLink ID="hlContact" runat="server" NavigateUrl="~/contact.aspx">Contact</asp:HyperLink></td>
        <td><asp:HyperLink ID="hlCredits" runat="server" NavigateUrl="~/credit.aspx">Credits</asp:HyperLink></td>
        <td><asp:HyperLink ID="hlSponsors" runat="server" NavigateUrl="~/sponsors.aspx">Sponsors</asp:HyperLink></td>
        <td><asp:HyperLink ID="hlPress" runat="server" NavigateUrl="~/Press.aspx">Press</asp:HyperLink></td>
    </tr> 
</table>
      </div>
</div>