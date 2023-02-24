<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="600" colspan="1" class="sHeader">On-Line Payment</td>
	</tr>
	<tr valign="top">
		<td width="450" colspan="1" height="1" bgcolor="#000000"></td>
	</tr>
	<tr class="sTitle" valign="top">
		
	</tr>
	<tr class="sTitle" valign="top">
	<%
		cid=Session("cid")
		strPid=Session("pid")
		intParcelNo=Session("varintParcelNo")
		strAddress=Session("varstrAddress")
		strName=Session("varstrName")
		intSect=Session("varintSect")
		intTwp=Session("varintTwp")
		intRange=Session("varintRange")

		payAmount = Session("payAmount")
		actionstr="Parcel.asp?pid=" & strPID & "&tid=0&cid=" & cid &"&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & ""
		Response.Redirect (actionstr)
	%>
	
	</tr>
	
		
</table>
