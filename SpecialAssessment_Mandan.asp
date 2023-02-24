<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="180" class="SHeader"></td>
		<td width="500" class="SHeader">City of Mandan Special Assessment Amounts</td>
		<td width="70" class="SHeader"></td>
		<td width="250" class="SHeader"></td>
	</tr>
	<tr valign="top">
		<td colspan="4" height="1" bgcolor="#000000"></td>
	</tr>
	<tr valign="top">
		<td class="STitle">Parcel</td>
		<td class="rText"><%= objRSPYGEN("TXPRCL") %></td>
		<td class="STitle">Name</td>
		<td class="rText"><%= objRSPYGEN("TPNAME") %></td>
	</tr>
	<%
	If objRSC10("ORIGCST") > 0  then
		Response.Write("<tr valign='top'>")
		Response.Write("<td class='STitle'>City Parcel #</td>")
		Response.Write("<td class='rText'>" & objRSC10("SEQ#") & " &nbsp;&nbsp;" & objRSC10("SPLIT") & "</td>")
		Response.Write("<td class='STitle'></td>")
		Response.Write("<td class='rText'></td>")
		Response.Write("</tr>")
	Else
		Response.Write("<tr valign='top'>")
		Response.Write("<td class='STitle'></td>")
		Response.Write("<td class='rText'></td>")
		Response.Write("<td class='STitle'></td>")
		Response.Write("<td class='rText'></td>")
		Response.Write("</tr>")
	End If
	%>
</table>
<br>
<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="90" class="SHeader">&nbsp;</td>
		<td width="4" class="SHeader"></td>
		<td width="40" class="SHeader"></td>
		<td width="4" class="SHeader"></td>
		<td width="40" class="SHeader"></td>
		<td width="4" class="SHeader"></td>
		<td width="140" class="SHeader"></td>
		<td width="4" class="SHeader"></td>
		<td width="40" class="SHeader"></td>
		<td width="4" class="SHeader"></td>
		<td width="40" class="SHeader"></td>
		<td width="4" class="SHeader"></td>
		<td width="40" class="SHeader"></td>
		<td width="4" class="SHeader"></td>
		<td width="40" class="SHeader"></td>
		<td width="4" class="SHeader"></td>
		<td width="40" class="SHeader"></td>
		<td width="4" class="SHeader"></td>
		<td width="40" class="SHeader"></td>
	</tr>
	<tr valign="top">
		<td colspan="15"></td>
	</tr>
		<tr valign="top">
		<td class="STitle">Type of Project</td>
		<td class="STitle"></td>
		<td class="STitle">Project Number</td>
		<td class="STitle"></td>
		<td align="center" class="STitle">Original Assessment</td>
		<td class="STitle"></td>
		<td align="center" class="STitle">Start Year</td>
		<td class="STitle"></td>
		<td align="center" align="right" class="STitle">Number of Years</td>
		<td class="STitle"></td>
		<td class="STitle">Interest Rate</td>
		<td class="STitle"></td>
		<td align="center" class="STitle">Remaining Balance</td>
		<td class="STitle"></td>
		<td align="center" class="STitle">Estimated Annual Principal</td>
		<td class="STitle"></td>
		<td align="center" class="STitle">Estimated Annual Interest</td>
		<td class="STitle"></td>
		<td align="center" align='right' class="STitle">Estimated Annual Installment</td>
		<td class="STitle"></td>
	</tr>
	<%


	if objRSC10("ORIGCST") <> 0 then
		While Not objRSCT10.EOF
			result = objRSCT10("REC#") MOD 2
			If result = 0 then
				Response.Write("<tr valign='top'>")
				If objRSCT10("TYPE") = "SIDEWALK-DRIVE-CURB-GUTTR" Then
					Response.Write("<td class='rText3'>SIDEWALK</td>")
				Else
					Response.Write("<td class='rText3'>"& objRSCT10("TYPE") & "</td>")
				End If
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='right' class='rText2'>" &  objRSCT10("DIST") & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='right' class='rText2'>" & FormatNumber( objRSCT10("OrigCst"), 2) & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='center' class='rText2'>" & objRSCT10("Year") & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='center' class='rText2'>" & objRSCT10("NoYear") & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='center' class='rText2'>" & FormatNumber(objRSCT10("IntRate"), 2) & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='right' class='rText2'>" & FormatNumber( objRSCT10("Balance"), 2) & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='right' class='rText2'>" &  calcZeroRS("YrlyPrin", 2) & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='right' class='rText2'>" &  calcZeroRS("YrlyInt", 2) & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='right' class='rText2'>" &  calcZeroRS("YrlyInst", 2) & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("</tr>")
			Else
				Response.Write("<tr valign='top'>")
				If objRSCT10("TYPE") = "SIDEWALK-DRIVE-CURB-GUTTR" Then
					Response.Write("<td class='rText1'>SIDEWALK</td>")
				Else
					Response.Write("<td class='rText1'>"& objRSCT10("TYPE") & "</td>")
				End If
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='right' class='rText'>" &  objRSCT10("DIST") & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='right' class='rText'>" & FormatNumber( objRSCT10("OrigCst"), 2) & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='center' class='rText'>" & objRSCT10("Year") & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='center' class='rText'>"& objRSCT10("NoYear") & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='center' class='rText'>" & FormatNumber(objRSCT10("IntRate"), 2) & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='right' class='rText'>" &  FormatNumber( objRSCT10("Balance"), 2) & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='right' class='rText'>" &  calcZeroRS("YrlyPrin", 2) & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='right' class='rText'>" &  calcZeroRS("YrlyInt", 2) & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("<td align='right' class='rText'>" &  calcZeroRS("YrlyInst", 2) & "</td>")
				Response.Write("<td class='rText'></td>")
				Response.Write("</tr>")
			End if
			objRSCT10.MoveNext
		Wend

		If result <> 0  Then
			Response.Write("<tr valign='top'>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td class='rText4' align='center' >Totals</td>")
			Response.Write("<td class='rText'></td>")
			Response.Write("<td align='right' class='rText2'>" & FormatNumber( objRSC10("Balance"), 2) & "</td>")
			Response.Write("<td class='rText'></td>")
			Response.Write("<td align='right' class='rText2'>" &  calcZeroRSC("YrlyPrin", 2) & "</td>")
			Response.Write("<td class='rText'></td>")
			Response.Write("<td align='right' class='rText2'>" &  calcZeroRSC("YrlyInt", 2) & "</td>")
			Response.Write("<td class='rText'></td>")
			Response.Write("<td align='right' class='rText2'>" &  calcZeroRSC("YrlyInst", 2) & "</td>")
			Response.Write("<td class='rText'></td>")
			Response.Write("</tr>")
		Else
			Response.Write("<tr valign='top'>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td></td>")
			Response.Write("<td class='rText4' align='center' >Totals</td>")
			Response.Write("<td class='rText'></td>")
			Response.Write("<td align='right' class='rText'>" & FormatNumber( objRSC10("Balance"), 2) & "</td>")
			Response.Write("<td class='rText'></td>")
			Response.Write("<td align='right' class='rText'>" &  calcZeroRSC("YrlyPrin", 2) & "</td>")
			Response.Write("<td class='rText'></td>")
			Response.Write("<td align='right' class='rText'>" &  calcZeroRSC("YrlyInt", 2) & "</td>")
			Response.Write("<td class='rText'></td>")
			Response.Write("<td align='right' class='rText'>" &  calcZeroRSC("YrlyInst", 2) & "</td>")
			Response.Write("<td class='rText'></td>")
			Response.Write("</tr>")
		End If

		Response.Write("<tr>")
		Response.Write("<td>")
		Response.Write("</td>")
		Response.Write("<td>")
		Response.Write("</td>")
		Response.Write("<td>")
		Response.Write("</td>")
		Response.Write("<td>")
		Response.Write("</td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
		Response.Write("<td colspan='40' ><B> Payoff figured to  " & objRSC10("TYPE") & " : $&nbsp;" & FormatNumber(objRSC10("OrigCst"), 2 ) & "<br>")
		Response.Write("Please call 701-667-3271 to confirm the payoff amount.<br>")
		Response.Write("Make checks payable to the City of Mandan</B></td>")
		Response.Write("</tr>")
	else
		Response.Write("<td class='rText2'>No Values</td>")
	End If

	Response.Write("<tr>")
	Response.Write("<td>")
	Response.Write("</td>")
	Response.Write("<td>")
	Response.Write("</td>")
	Response.Write("<td>")
	Response.Write("</td>")
	Response.Write("<td>")
	Response.Write("</td>")
	Response.Write("</tr>")
	%>
</table>