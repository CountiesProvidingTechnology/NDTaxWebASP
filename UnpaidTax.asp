<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="600" colspan="8" class="sHeader">Unpaid Taxes</td>
	</tr>
	<tr valign="top">
		<td width="450" colspan="8" height="1" bgcolor="#000000"></td>
	</tr>
	<tr valign="top">
		<td width="40" class="sTitle">Year</td>
		<td width="80" class="sTitle" align="right">Net Tax</td>
		<td width="80" class="sTitle" align="right">DIS/PN/IN</td>
		<td width="80" class="sTitle" align="right">Special Asmt</td>
		<td width="80" class="sTitle" align="right">Special Asmt Penalty</td>
		<td width="80" class="sTitle" align="right">Advertising</td>
		<td width="80" class="sTitle" align="right">Total Due</td>
		<td width="80" class="sTitle" align="right"></td>
	</tr>
	<tr valign="top">
<%
If not objRS5.eof then

	If objRS5("UPYEAR1") = 0 Then
		Response.Write("<td width='80' class='rText' colspan='7'>No Unpaid</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
	Else
		Response.Write("<td width='40' class='rText'>" & objRS5("UPYEAR1") & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPTAX1"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPPEN1"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPSPEC1"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPSAPEN1"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPADV1"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPTOTDUE1"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>As of</td>")
		Response.Write("</tr>")
	If objRS5("UPYEAR2") > 0 Then
		Response.Write("<tr valign='top'>")
		Response.Write("<td width='40' class='rText'>" & objRS5("UPYEAR2") & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPTAX2"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPPEN2"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPSPEC2"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPSAPEN2"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPADV2"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPTOTDUE2"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcDateRS5(("UPINDT")) & "</td>")
		Response.Write("</tr>")
	else
		'Response.Write(" the second record year here now " )
		Response.Write("<tr valign='top'>")
		Response.Write("<td width='40' class='rText'></td>")
		Response.Write("<td width='80' class='rText' align='right'></td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcDateRS5(("UPINDT")) &  "</td>")
		Response.Write("</tr>")
	End If
	If objRS5("UPYEAR3") > 0 Then
		Response.Write("<tr valign='top'>")
		Response.Write("<td width='40' class='rText'>" & objRS5("UPYEAR3") & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPTAX3"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPPEN3"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPSPEC3"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPSAPEN3"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPADV3"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPTOTDUE3"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'></td>")
		Response.Write("</tr>")
	End If
	If objRS5("UPYEAR4") > 0 Then
		Response.Write("<tr valign='top'>")
		Response.Write("<td width='40' class='rText'>" & objRS5("UPYEAR4") & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPTAX4"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPPEN4"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPSPEC4"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPSAPEN4"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPADV4"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPTOTDUE4"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'></td>")
		Response.Write("</tr>")
	End If
	If objRS5("UPYEAR5") > 0 Then
		Response.Write("<tr valign='top'>")
		Response.Write("<td width='40' class='rText'>" & objRS5("UPYEAR5") & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPTAX5"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPPEN5"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPSPEC5"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPSAPEN5"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPADV5"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS5(("UPTOTDUE5"), 2) & "</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("</tr>")

		Response.Write("<tr valign='top'>")
		Response.Write("<td width='40' class='rText'></td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>Total Due</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(calcTotRS5(calcTot), 2) & "</td>")
		Response.Write("</tr>")
	else
		Response.Write("<tr valign='top'>")
		Response.Write("<td width='40' class='rText'></td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>Total Due</td>")
		Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(calcTotRS5(calcTot), 2) & "</td>")
		Response.Write("</tr>")
	End If

		Response.Write("<tr valign='top'>")
		Response.Write("<td width='40' class='rText'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
		Response.Write("</tr>")
	End If
End If

	If not objRS52.eof then
		If objRS52("UPYEAR1") = 0 Then
			'Response.Write(" the second record here now " )
			If objRS52("UPYEAR1") = 0 Then
			else
			Response.Write("<td width='80' class='rText' colspan='8'>No Unpaid</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			End If
		Else
			Response.Write("<tr valign='top'>")
			Response.Write("<td width='40' class='rText'>" & objRS52("UPYEAR1") & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPTAX1"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPPEN1"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPSPEC1"), 2) & strData & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPSAPEN1"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPADV1"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPTOTDUE1"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>As of</td>")
			Response.Write("</tr>")
		End If
		If objRS52("UPYEAR2") > 0 Then
			Response.Write("<tr valign='top'>")
			Response.Write("<td width='40' class='rText'>" & objRS52("UPYEAR2") & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPTAX2"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPPEN2"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPSPEC2"), 2) & strData & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPSAPEN2"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPADV2"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPTOTDUE2"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcDateMoYrRS52(("UPINDT")) & "</td>")
			Response.Write("</tr>")
		else
			Response.Write("<tr valign='top'>")
			Response.Write("<td width='40' class='rText'></td>")
			Response.Write("<td width='80' class='rText' align='right'></td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			If objRS5("UPYEAR1") = 0 Then
			Response.Write("<td width='80' class='rText' align='right'></td>")
			else
			Response.Write("<td width='80' class='rText' align='right'>" & calcDateMoYrRS52(("UPINDT")) &  "</td>")
			end if
			Response.Write("</tr>")
		End If
		If objRS52("UPYEAR3") > 0 Then
			Response.Write("<tr valign='top'>")
			Response.Write("<td width='40' class='rText'>" & objRS52("UPYEAR3") & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS52("UPTAX3"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS52("UPPEN3"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPSPEC3"), 2) & strData & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPSAPEN3"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPADV3"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPTOTDUE3"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & strData & "</td>")
			Response.Write("</tr>")
		End If
		If objRS52("UPYEAR4") > 0 Then
			Response.Write("<tr valign='top'>")
			Response.Write("<td width='40' class='rText'>" & objRS52("UPYEAR4") & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS52("UPTAX4"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS52("UPPEN4"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPSPEC4"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPSAPEN4"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPADV4"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPTOTDUE4"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'></td>")
			Response.Write("</tr>")
		End If
		If objRS52("UPYEAR5") > 0 Then
			Response.Write("<tr valign='top'>")
			Response.Write("<td width='40' class='rText'>" & objRS52("UPYEAR5") & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS52("UPTAX5"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS52("UPPEN5"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPSPEC5"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPSAPEN5"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPADV5"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>" & calcZeroRS52(("UPTOTDUE5"), 2) & "</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("</tr>")

			Response.Write("<tr valign='top'>")
			Response.Write("<td width='40' class='rText'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>Total Due </td>")
			Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(calcTotRS52(calcTot2), 2) & "</td>")
			Response.Write("</tr>")
		else
			If objRS52("UPYEAR1") = 0 Then
			else
			Response.Write("<tr valign='top'>")
			Response.Write("<td width='40' class='rText'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("<td width='80' class='rText' align='right'>Total Due </td>")
			Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(calcTotRS52(calcTot2), 2) & "</td>")
			Response.Write("</tr>")
			End If
		End If
	End If
%>
</table>