

<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="600" colspan="9" class="sHeader">5 Year Tax History</td>
	</tr>
	<tr valign="top">
		<td width="450" colspan="9" height="1" bgcolor="#000000"></td>
	</tr>
	<tr valign="top">
		<td width="40" class="sTitle">Year</td>
		<td width="80" class="sTitle" align="right">Est/Tax<br>Market</td>
		<td width="80" class="sTitle" align="right">TC<br>Value</td>
		<td width="80" class="sTitle" align="right">TC/Mkt<br>Rate</td>
		<td width="80" class="sTitle" align="right">Other<br>Credit</td>
		<td width="80" class="sTitle" align="right">Abat/<br>Add</td>
		<td width="80" class="sTitle" align="right">Special<br>Asmts</td>
		<td width="80" class="sTitle" align="right">Total<br>Net Tax</td>
	</tr>
	<tr valign="top">
		<td class="rText" height="25"><%= objRSPY("PHYR01") %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHEST1"), 0) %><br><%= calcZeroPY(("PHTMV1"), 0) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHTCV1"), 0) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHTCR1"), 5) %><br><%= calcZeroPY(("PHMTR1"), 5) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHOCR1"), 2) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHABT1"), 2) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHSPC1"), 2) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHNET1"), 2) %></td>
	</tr>
	<tr valign="top">
		<td class="rText2" height="25"><%= objRSPY("PHYR02") %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHEST2"), 0) %><br><%= calcZeroPY(("PHTMV2"), 0) %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHTCV2"), 0) %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHTCR2"), 5) %><br><%= calcZeroPY(("PHMTR2"), 5) %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHOCR2"), 2) %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHABT2"), 2) %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHSPC2"), 2) %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHNET2"), 2) %></td>
	</tr>
	<tr valign="top">
		<td class="rText" height="25"><%= objRSPY("PHYR03") %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHEST3"), 0) %><br><%= calcZeroPY(("PHTMV3"), 0) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHTCV3"), 0) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHTCR3"), 5) %><br><%= calcZeroPY(("PHMTR3"), 5) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHOCR3"), 2) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHABT3"), 2) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHSPC3"), 2) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHNET3"), 2) %></td>
	</tr>
	<tr valign="top">
		<td class="rText2" height="25"><%= objRSPY("PHYR04") %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHEST4"), 0) %><br><%= calcZeroPY(("PHTMV4"), 0) %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHTCV4"), 0) %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHTCR4"), 5) %><br><%= calcZeroPY(("PHMTR4"), 5) %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHOCR4"), 2) %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHABT4"), 2) %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHSPC4"), 2) %></td>
		<td class="rText2" align="right"><%= calcZeroPY(("PHNET4"), 2) %></td>
	</tr>
	<tr valign="top">
		<td class="rText" height="25"><%= objRSPY("PHYR05") %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHEST5"), 0) %><br><%= calcZeroPY(("PHTMV5"), 0) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHTCV5"), 0) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHTCR5"), 5) %><br><%= calcZeroPY(("PHMTR5"), 5) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHOCR5"), 2) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHABT5"), 2) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHSPC5"), 2) %></td>
		<td class="rText" align="right"><%= calcZeroPY(("PHNET5"), 2) %></td>
	</tr>
</table>
