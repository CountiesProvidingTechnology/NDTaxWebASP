

<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="80" class="SHeader"></td>
		<td width="200" class="SHeader">Ditch Assessment Amounts</td>
		<td width="70" class="SHeader"></td>
		<td width="250" class="SHeader"></td>
	</tr>
	<tr valign="top">
		<td colspan="4" height="1" bgcolor="#000000"></td>
	</tr>
	<tr valign="top">
		<td class="STitle">Parcel</td>
		<td class="rText"><%= objRSSP("TXPRCL") %></td>
		<td class="STitle">Name</td>
		<td class="rText"><%= objRS("TXTNAM") %></td>
	</tr>
	<tr>
		<td class="STitle"></td>
		<td class="rText"></td>
		<td class="STitle"></td>
		<td class="rText"></td>
	</tr>
</table>
<br><br>
<table border="0" cellpadding="0" cellspacing="0">
  <tr valign="top">
    <td width="40" class="SHeader">&nbsp;</td>
    <td width="4" class="SHeader"></td>
    <td width="40" class="SHeader"></td>
    <td width="4" class="SHeader"></td>
    <td width="40" class="SHeader"></td>
    <td width="4" class="SHeader"></td>
    <td width="10" class="SHeader"></td>
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
    <td class="STitle" >Type Code</td>
    <td class="STitle"></td>
    <td class="STitle">Record</td>
    <td class="STitle"></td>
    <td class="STitle">Description</td>
    <td class="STitle"></td>
    <td align="right" class="STitle" nowrap>Original Benefit</td>
    <td class="STitle"></td>
    <td align="right" class="STitle">Paid Up</td>
  </tr>
   <%
    For recnum = 1 To 10
    	if Trim(objRSSP("DSSA#" & recnum)) = 0 Then
    	if recnum < 2 then
    	Response.Write("<td class='rText2'nowrap> No values </td>")
    	end if
    		Exit For
    	else
    	result = recnum MOD 2
    	If result = 0 then
        Response.Write("<tr valign='top'>")
  	  Response.Write("<td class='rText2'>" &  objRSSP("DSSAC" & recnum) & "</td>")
  	  Response.Write("<td class='rText2'></td>")
  	  Response.Write("<td class='rText2'>" & objRSSP("DSSA#" & recnum) & "</td>")
  	  Response.Write("<td class='rText2'>&nbsp;&nbsp;</td>")
  	  Response.Write("<td align='right' class='rText2'nowrap>" &  objRSSP("DSSAD" & recnum) & "</td>")
  	  Response.Write("<td class='rText2'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & FormatNumber(objRSSP(("DSSAA" & recnum)), 2) & "</td>")
  	  Response.Write("<td class='rText2'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & objRSSP("DSSAF" & recnum) & "</td>")
  	  Response.Write("</tr>")
  	Else
        Response.Write("<tr valign='top'>")
  	  Response.Write("<td class='rText'>" & objRSSP("DSSAC" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td class='rText'>" & objRSSP("DSSA#" & recnum) & "</td>")
  	  Response.Write("<td class='rText'>&nbsp;&nbsp;</td>")
  	  Response.Write("<td align='right' class='rText'nowrap>" &  objRSSP("DSSAD" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" & FormatNumber(objRSSP(("DSSAA" & recnum)), 2) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" & objRSSP("DSSAF" & recnum) & "</td>")
  	  Response.Write("</tr>")
  	End if
  	End if
    Next
  %>

</table>

