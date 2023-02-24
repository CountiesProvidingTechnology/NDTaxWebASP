<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="600" colspan="9" class="sHeader">5 Year Tax History</td>
	</tr>
	<tr valign="top">
		<td width="450" colspan="9" height="1" bgcolor="#000000"></td>
	</tr>
	<tr valign="top">
		<td width="40" class="sTitle">Year</td>
		<td width="80" class="sTitle" align="right">Est<br>Assessed</td>
		<td width="80" class="sTitle" align="right">Taxable<br>Net Taxable</td>
		<td width="80" class="sTitle" align="right">Mill Rate</td>
		<td width="80" class="sTitle" align="right">Abate/  Added</td>
		<td width="80" class="sTitle" align="right">Special <br>Asmts</td>
		<td width="80" class="sTitle" align="right">Total <br>Tax</td>
		<td width="80" class="sTitle" align="right"></td>
	</tr>
	<tr valign="top">
		<td class="rText" height="25"><%= objRS("RCYR1") %></td>
		<td class="rText" align="right"><%= FormatNumber(objRS("RCEST1"), 0) %><br><%= FormatNumber(objRS("RCGROSS1"), 0) %></td>
		<td class="rText" align="right"><%= FormatNumber(objRS("RCNET1"), 0) %><br><%= FormatNumber(objRS("RCTAX1"), 0) %></td>
		<td class="rText" align="right"><%= objRS("RCRATE1") %></td>
		<td class="rText" align="right"><%= objRS("RCABAD1") %></td>
		<td class="rText" align="right"><%= FormatNumber(objRS("RCSA1"), 2) %></td>
		<td class="rText" align="right"><%= FormatNumber(objRS("RCNETTAX1"), 2) %></td>
		<td class="rText" align="right"></td>
	</tr>
	<tr valign="top">
		<td class="rText2" height="25"><%= objRS("RCYR2") %></td>
		<td class="rText2" align="right"><%= FormatNumber(objRS("RCEST2"), 0) %><br><%= FormatNumber(objRS("RCGROSS2"), 0) %></td>
		<td class="rText2" align="right"><%= FormatNumber(objRS("RCNET2"), 0) %><br><%= FormatNumber(objRS("RCTAX2"), 0) %></td>
		<td class="rText2" align="right"><%= objRS("RCRATE2") %></td>
		<td class="rText2" align="right"><%= objRS("RCABAD2") %></td>
		<td class="rText2" align="right"><%= FormatNumber(objRS("RCSA2"), 2) %></td>
		<td class="rText2" align="right"><%= FormatNumber(objRS("RCNETTAX2"), 2) %></td>
		<td class="rText2" align="right"></td>
	</tr>
	<tr valign="top">
		<td class="rText" height="25"><%= objRS("RCYR3") %></td>
		<td class="rText" align="right"><%= FormatNumber(objRS("RCEST3"), 0) %><br><%= FormatNumber(objRS("RCGROSS3"), 0) %></td>
		<td class="rText" align="right"><%= FormatNumber(objRS("RCNET3"), 0) %><br><%= FormatNumber(objRS("RCTAX3"), 0) %></td>
		<td class="rText" align="right"><%= objRS("RCRATE3") %>
		<td class="rText" align="right"><%= objRS("RCABAD3") %></td>
		<td class="rText" align="right"><%= FormatNumber(objRS("RCSA3"), 2) %></td>
		<td class="rText" align="right"><%= FormatNumber(objRS("RCNETTAX3"), 2) %></td>
		<td class="rText" align="right"></td>
	</tr>
	<tr valign="top">
		<td class="rText2" height="25"><%= objRS("RCYR4") %></td>
		<td class="rText2" align="right"><%= FormatNumber(objRS("RCEST4"), 0) %><br><%= FormatNumber(objRS("RCGROSS4"), 0) %></td>
		<td class="rText2" align="right"><%= FormatNumber(objRS("RCNET4"), 0) %><br><%= FormatNumber(objRS("RCTAX4"), 0) %></td>
		<td class="rText2" align="right"><%= objRS("RCRATE4") %>
		<td class="rText2" align="right"><%= objRS("RCABAD4") %></td>
		<td class="rText2" align="right"><%= FormatNumber(objRS("RCSA4"), 2) %></td>
		<td class="rText2" align="right"><%= FormatNumber(objRS("RCNETTAX4"), 2) %></td>
		<td class="rText2" align="right"></td>
	</tr>
	<tr valign="top">
		<td class="rText" height="25"><%= objRS("RCYR5") %></td>
		<td class="rText" align="right"><%= FormatNumber(objRS("RCEST5"), 0) %><br><%= FormatNumber(objRS("RCGROSS5"), 0) %></td>
		<td class="rText" align="right"><%= FormatNumber(objRS("RCNET5"), 0) %><br><%= FormatNumber(objRS("RCTAX5"), 0) %></td>
		<td class="rText" align="right"><%= objRS("RCRATE5") %>
		<td class="rText" align="right"><%= objRS("RCABAD5") %></td>
		<td class="rText" align="right"><%= FormatNumber(objRS("RCSA5"), 2) %></td>
		<td class="rText" align="right"><%= FormatNumber(objRS("RCNETTAX5"), 2) %></td>
		<td class="rText" align="right"></td>
	</tr>
</table>