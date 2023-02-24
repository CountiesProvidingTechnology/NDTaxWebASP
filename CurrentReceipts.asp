<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="600" colspan="6" class="sHeader">Tax Receipt Information</td>
	</tr>
	<tr valign="top">
		<td width="450" colspan="6" height="1" bgcolor="#000000"></td>
	</tr>
	<tr valign="top">
		<td width="280" class="sTitle" colspan="2"></td>
		<td width="80" class="sTitle" align="right">Tax &amp;</td>
		<td width="240" class="sTitle" align="center" colspan="3">Special Assessments</td>
	</tr>
	<tr valign="top">
		<td width="220" class="sTitle"></td>
		<td width="60" class="sTitle" align="right"></td>
		<td width="80" class="sTitle" align="right">Penalty</td>
		<td width="80" class="sTitle" align="right">Rec #</td>
		<td width="80" class="sTitle" align="right">Code</td>
		<td width="80" class="sTitle" align="right">Amount</td>
	</tr>
	<%=	printTaxRecord() %>
	<%
	if Session("recnumberend") = 5 then
	%>
		<%= printTaxRecord4() %>
	<%
		end if
	%>
		<%
		if Session("recnumberend") = 10 then
		%>
			<%= printTaxRecord6() %>
		<%
			end if
		%>
		<%
				if Session("recnumberend") = 15 then
				%>
					<%= printTaxRecord7() %>
				<%
					end if
		%>
	<%'='	printTaxRecord4() %>
	<tr valign="top">
		<td width="10"></td>
		<td width="10"></td>
		<td width="10"></td>
		<td width="10"></td>
		<td width="10"></td>

</tr>
</table>