<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="Merchandise.aspx.vb" Inherits="Merchandise"  %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<table ID="Table1">
			<TBODY>
				<tr>
					<td colspan="2"><h1>
							Merchandise</h1>
					</td>
				</tr>
				<tr>
					<td width="200"></td>
					<td width="500">
						<table width="100%" border="0" cellpadding="5" cellspacing="5">
							<TR>
								<TD></TD>
								<TD></TD>
								<TD>
									<form target="paypal" action="https://www.paypal.com/cgi-bin/webscr" method="post" ID="Form2">
										<input type="hidden" name="cmd" value="_cart" ID="Hidden1"> <input type="hidden" name="business" value="info@vivalesamis.com" ID="Hidden2">
										<input type="image" src="https://www.paypal.com/en_US/i/btn/view_cart_02.gif" border="0"
											name="submit" alt="Make payments with PayPal - it's fast, free and secure!" ID="Image1">
										<input type="hidden" name="display" value="1" ID="Hidden3">
									</form>
								</TD>
							</TR>
							<TR>
								<TD>
									<DIV align="center"><A href="javascript:openPopup('images/dvdcover.jpg','DVDWin',630,443);"><IMG src="dvd/dvdcover.jpg" width="150" border="0"></A></DIV>
								</TD>
								<TD>
									<B>Viva Les Amis DVD</B>
									<BR>
									The Viva Les Amis DVD includes:<BR>
									-the documentary - 49 minutes
									<BR>
									-a slide show of photos<BR>
									-group photos<BR>
									-Excerpts from Newman's "Tell me a Joke" movie
									<BR>
									$20.00
								</TD>
								<TD>
									<FORM id="Form7" action="https://www.paypal.com/cgi-bin/webscr" method="post" target="paypal">
										<INPUT id="Image7" type="image" alt="Make payments with PayPal - it's fast, free and secure!"
											src="https://www.paypal.com/en_US/i/btn/x-click-but22.gif" border="0" name="submit">
										<IMG height="1" alt="" src="https://www.paypal.com/en_US/i/scr/pixel.gif" width="1" border="0">
										<INPUT id="Hidden71" type="hidden" value="1" name="add"> <INPUT id="Hidden72" type="hidden" value="_cart" name="cmd">
										<INPUT id="Hidden73" type="hidden" value="info@vivalesamis.com" name="business">
										<INPUT id="Hidden74" type="hidden" value="Viva Les Amis DVD" name="item_name"> <INPUT id="Hidden75" type="hidden" value="VLD-001" name="item_number">
										<INPUT id="Hidden76" type="hidden" value="20.00" name="amount"> <INPUT id="Hidden7" type="hidden" value="Primary" name="page_style">
										<INPUT id="Hidden8" type="hidden" value="2" name="no_shipping"> <INPUT id="Hidden9" type="hidden" value="http://vivalesamis.com/thankyou.aspx" name="return">
										<INPUT id="Hidden10" type="hidden" value="http://vivalesamis.com/merchandiseFS.htm" name="cancel_return">
										<INPUT id="Hidden11" type="hidden" value="Add any comments here" name="cn"> <INPUT id="Hidden12" type="hidden" value="USD" name="currency_code">
										<INPUT id="Hidden13" type="hidden" value="US" name="lc"> <INPUT id="Hidden14" type="hidden" value="PP-ShopCartBF" name="bn">
									</FORM>
								</TD>
							</TR>
							<TR>
								<TD><A href="javascript:openPopup('images/girltanktop.jpg','DVDWin',630,443);"><IMG src="images/girltanktop.jpg" width="150" border="0"></A></TD>
								<TD><B>Black Tank Top</B>
									<BR>
									This black tank top is an American Apparel standard.100% cotton. Good fit. 
									$15.00</TD>
								<TD>
									<FORM id="Form1" action="https://www.paypal.com/cgi-bin/webscr" method="post" target="paypal">
										<TABLE id="Table2">
											<TR>
												<TD><INPUT id="Hidden15" type="hidden" value="Size" name="on0">Size</TD>
												<TD><SELECT id="Select1" name="os0">
												<option value="Small" selected>Small</option>
												l<OPTION value="Medium" >Medium</option>
													<OPTION value="Large">Large</OPTION>
													</SELECT>
												</TD>
											</TR>
										</TABLE>
										<INPUT id="Image2" type="image" alt="Make payments with PayPal - it's fast, free and secure!"
											src="https://www.paypal.com/en_US/i/btn/x-click-but22.gif" border="0" name="submit">
										<IMG height="1" alt="" src="https://www.paypal.com/en_US/i/scr/pixel.gif" width="1" border="0">
										<INPUT id="Hidden16" type="hidden" value="1" name="add"> <INPUT id="Hidden17" type="hidden" value="_cart" name="cmd">
										<INPUT id="Hidden18" type="hidden" value="info@vivalesamis.com" name="business">
										<INPUT id="Hidden19" type="hidden" value="Girls Black Tank Top" name="item_name">
										<INPUT id="Hidden20" type="hidden" value="VLT-001" name="item_number"> <INPUT id="Hidden21" type="hidden" value="15.00" name="amount">
										<INPUT id="Hidden22" type="hidden" value="Primary" name="page_style"> <INPUT id="Hidden23" type="hidden" value="2" name="no_shipping">
										<INPUT id="Hidden24" type="hidden" value="http://vivalesamis.com/thankyou.aspx" name="return">
										<INPUT id="Hidden25" type="hidden" value="http://vivalesamis.com/merchandiseFS.htm" name="cancel_return">
										<INPUT id="Hidden26" type="hidden" value="USD" name="currency_code"> <INPUT id="Hidden27" type="hidden" value="US" name="lc">
										<INPUT id="Hidden28" type="hidden" value="PP-ShopCartBF" name="bn">
									</FORM>
								</TD>
							</TR>
							<TR>
								<TD><A href="javascript:openPopup('images/menshirtblack.jpg','DVDWin',630,443);"><IMG src="images/menshirtblack.jpg" width="150" border="0"></A></TD>
								<TD><B>Men's Black Shirt</B>
									<BR>
									This 100% cotton, American Apparel shirt is good quality and extra soft. Women 
									might want to wear this one, too. $15.00
								</TD>
								<TD>
									<FORM id="Form4" action="https://www.paypal.com/cgi-bin/webscr" method="post" target="paypal">
										<TABLE id="Table4">
											<TR>
												<TD><INPUT id="Hidden43" type="hidden" value="Size" name="on0">Size</TD>
												<TD><SELECT id="Select3" name="os0"><OPTION value="Small" selected>
														Small<OPTION value="Medium">
														Medium<OPTION value="Large">Large</OPTION>
														<OPTION value="X-Large">X-Large</OPTION>
													</SELECT>
												</TD>
											</TR>
										</TABLE>
										<INPUT id="Image4" type="image" alt="Make payments with PayPal - it's fast, free and secure!"
											src="https://www.paypal.com/en_US/i/btn/x-click-but22.gif" border="0" name="submit">
										<IMG height="1" alt="" src="https://www.paypal.com/en_US/i/scr/pixel.gif" width="1" border="0">
										<INPUT id="Hidden44" type="hidden" value="1" name="add"> <INPUT id="Hidden45" type="hidden" value="_cart" name="cmd">
										<INPUT id="Hidden46" type="hidden" value="info@vivalesamis.com" name="business">
										<INPUT id="Hidden47" type="hidden" value="Men's Black Tee" name="item_name"> <INPUT id="Hidden48" type="hidden" value="VLT-003" name="item_number">
										<INPUT id="Hidden49" type="hidden" value="15.00" name="amount"> <INPUT id="Hidden50" type="hidden" value="Primary" name="page_style">
										<INPUT id="Hidden51" type="hidden" value="2" name="no_shipping"> <INPUT id="Hidden52" type="hidden" value="http://vivalesamis.com/thankyou.aspx" name="return">
										<INPUT id="Hidden53" type="hidden" value="http://vivalesamis.com/merchandiseFS.htm" name="cancel_return">
										<INPUT id="Hidden54" type="hidden" value="USD" name="currency_code"> <INPUT id="Hidden55" type="hidden" value="US" name="lc">
										<INPUT id="Hidden56" type="hidden" value="PP-ShopCartBF" name="bn">
									</FORM>
								</TD>
							</TR>
							<TR>
								<TD><A href="javascript:openPopup('images/menshirtmaroon.jpg','DVDWin',630,443);"><IMG src="images/menshirtmaroon.jpg" width="150" border="0"></A></TD>
								<TD><B>Men's Maroon Shirt</B>
									<BR>
									This 100% cotton, American Apparel shirt is the same style as the Men's Black 
									Shirt, but in the color of the original Les Amis awning. $15.00
								</TD>
								<TD>
									<FORM id="Form5" action="https://www.paypal.com/cgi-bin/webscr" method="post" target="paypal">
										<TABLE id="Table5">
											<TR>
												<TD><INPUT id="Hidden57" type="hidden" value="Size" name="on0">Size</TD>
												<TD><SELECT id="Select4" name="os0">
														<OPTION value="Small" selected>
														Small<OPTION value="Medium">
														Medium<OPTION value="Large">Large
														<option value="XLarge">XLarge
														</OPTION>
													</SELECT>
												</TD>
											</TR>
										</TABLE>
										<INPUT id="Image5" type="image" alt="Make payments with PayPal - it's fast, free and secure!"
											src="https://www.paypal.com/en_US/i/btn/x-click-but22.gif" border="0" name="submit">
										<IMG height="1" alt="" src="https://www.paypal.com/en_US/i/scr/pixel.gif" width="1" border="0">
										<INPUT id="Hidden58" type="hidden" value="1" name="add"> <INPUT id="Hidden59" type="hidden" value="_cart" name="cmd">
										<INPUT id="Hidden60" type="hidden" value="info@vivalesamis.com" name="business">
										<INPUT id="Hidden61" type="hidden" value="Men's Maroon Shirt" name="item_name"> <INPUT id="Hidden62" type="hidden" value="VLT-003" name="item_number">
										<INPUT id="Hidden63" type="hidden" value="15.00" name="amount"> <INPUT id="Hidden64" type="hidden" value="Primary" name="page_style">
										<INPUT id="Hidden65" type="hidden" value="2" name="no_shipping"> <INPUT id="Hidden66" type="hidden" value="http://vivalesamis.com/thankyou.aspx" name="return">
										<INPUT id="Hidden67" type="hidden" value="http://vivalesamis.com/merchandiseFS.htm" name="cancel_return">
										<INPUT id="Hidden68" type="hidden" value="USD" name="currency_code"> <INPUT id="Hidden69" type="hidden" value="US" name="lc">
										<INPUT id="Hidden70" type="hidden" value="PP-ShopCartBF" name="bn">
									</FORM>
								</TD>
							</TR>
							<TR>
								<TD></TD>
								<TD></TD>
								<TD>
									<form target="paypal" action="https://www.paypal.com/cgi-bin/webscr" method="post" ID="Form6">
										<input type="hidden" name="cmd" value="_cart" ID="Hidden4"> <input type="hidden" name="business" value="info@vivalesamis.com" ID="Hidden5">
										<input type="image" src="https://www.paypal.com/en_US/i/btn/view_cart_02.gif" border="0"
											name="submit" alt="Make payments with PayPal - it's fast, free and secure!" ID="Image6">
										<input type="hidden" name="display" value="1" ID="Hidden6">
									</form>
								</TD>
							</TR>
						</table>
					</td>
				</tr>
			</TBODY>
		</table>
		<br /><br /><br />
</asp:Content>

