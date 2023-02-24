

<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="80" class="SHeader"></td>
		<td width="200" class="SHeader">Sales Ratio</td>
		<td width="70" class="SHeader"></td>
		<td width="250" class="SHeader"></td>
	</tr>
	<tr valign="top">
		<td colspan="4" height="1" bgcolor="#000000"></td>
	</tr>
	<tr valign="top">
		<td class="STitle">Parcel</td>
		<td class="rText"><%= objRS("TXPRCL") %></td>
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
    <td class="STitle"nowrap>CRV# |</td>
    <td class="STitle"></td>
    <td class="STitle"nowrap>Buyer </td>
    <td class="STitle"></td>
    <td class="STitle"nowrap>Seller </td>
    <td class="STitle"></td>
    <td class="STitle"nowrap>Sale Price |</td>
    <td class="STitle"></td>
    <td class="STitle"nowrap>Mrkt Value |</td>
    <td class="STitle"></td>
    <td class="STitle"nowrap>Adj |</td>
    <td class="STitle"></td>
    <td class="STitle"nowrap>Sale Date |</td>
    <td class="STitle"></td>
    <td class="STitle"nowrap>Trnsc|</td>
    <td class="STitle"></td>
    <td class="STitle"nowrap>Sale Desc|</td>
    <td class="STitle"></td>
    <td class="STitle"nowrap>PropType|</td>
    <td class="STitle"></td>
    <td class="STitle"nowrap>Aud Date</td>
    <td class="STitle"></td>

  </tr>
  <%
  recnum = 1
If not objRS10.eof then

  For recnum = 1 To 10
'Response.Write("the value of the recnum: "  & recnum  )
  	if Trim(objRS10("SRAUDT" & recnum)) = 0  Then
  		'recnum = 10
  		'Response.Write("the value of the recnum: "  & recnum  )
  		If recnum < 2 then
  		Response.Write("<td class='rText2'>No Values</td>")
  		end if
  		Exit For
  	else
  	'Response.Write("the value of the var: " & objRS10("SRPRCH" & recnum) )
  	result = recnum MOD 2
  	'Response.Write("the value of recnum and result is : " & recnum & result)
  	If result = 0 then
      Response.Write("<tr valign='top'>")
	  Response.Write("<td class='rText2'>" &  objRS10("SRCRV" & recnum) & "</td>")
	  Response.Write("<td class='rText'></td>")
	  Response.Write("<td class='rText2'>" & objRS10("SRBNAM" & recnum) & "</td>")
	  Response.Write("<td class='rText'>&nbsp;</td>")
	  Response.Write("<td class='rText2'>" & objRS10("SRSNAM" & recnum) & "</td>")
	  Response.Write("<td class='rText'>&nbsp;</td>")
	  Response.Write("<td align='right' class='rText2'>" & objRS10("SRPRCH" & recnum) & "</td>")
	  Response.Write("<td class='rText'>&nbsp;</td>")
	  Response.Write("<td align='right' class='rText2'>" & objRS10("SRMVCR" & recnum) & "</td>")
	  Response.Write("<td class='rText'>&nbsp;</td>")
	  Response.Write("<td align='right' class='rText2'>" & objRS10("SRAJPR" & recnum) & "</td>")
	  Response.Write("<td class='rText'>&nbsp;</td>")
	  Response.Write("<td align='right' class='rText2'>" & calcDate6Digit(("SRSLDT"), recnum) & "</td>")
	  Response.Write("<td class='rText'>&nbsp;</td>")
	  Response.Write("<td align='right' class='rText2'>" & objRS10("TRAND" & recnum) & "</td>")
	  Response.Write("<td class='rText'>&nbsp;</td>")
	  Response.Write("<td class='rText2'>" & objRS10("SLCDD" & recnum) & "</td>")
	  Response.Write("<td class='rText'>&nbsp;</td>")
	  Response.Write("<td class='rText2'>" & objRS10("PRPD" & recnum) & "</td>")
	  Response.Write("<td class='rText'>&nbsp;</td>")
	  Response.Write("<td align='right' class='rText2'>" & calcDateRS10(("SRAUDT"), recnum) & "</td>")
	  Response.Write("<td class='rText'>&nbsp;</td>")
	  Response.Write("</tr>")
	Else
      Response.Write("<tr valign='top'>")
	  Response.Write("<td class='rText'>" &  objRS10("SRCRV" & recnum) & "</td>")
	  Response.Write("<td class='rText2'>&nbsp;</td>")
	  Response.Write("<td class='rText'>" & objRS10("SRBNAM" & recnum) & "</td>")
	  Response.Write("<td class='rText2'>&nbsp;</td>")
	  Response.Write("<td class='rText'>" & objRS10("SRSNAM" & recnum) & "</td>")
	  Response.Write("<td class='rText2'>&nbsp;</td>")
	  Response.Write("<td align='right' class='rText'>" & objRS10("SRPRCH" & recnum) & "</td>")
	  Response.Write("<td class='rText2'>&nbsp;</td>")
	  Response.Write("<td align='right' class='rText'>" & objRS10("SRMVCR" & recnum) & "</td>")
	  Response.Write("<td class='rText2'>&nbsp;</td>")
	  Response.Write("<td align='right' class='rText'>" & objRS10("SRAJPR" & recnum) & "</td>")
	  Response.Write("<td class='rText2'>&nbsp;</td>")
	  Response.Write("<td align='right' class='rText'>" & calcDate6Digit(("SRSLDT"), recnum) & "</td>")
	  Response.Write("<td class='rText2'>&nbsp;</td>")
	  Response.Write("<td align='right' class='rText'>" & objRS10("TRAND" & recnum) & "</td>")
	  Response.Write("<td class='rText2'>&nbsp;</td>")
	  Response.Write("<td class='rText'>" & objRS10("SLCDD" & recnum) & "</td>")
	  Response.Write("<td class='rText2'>&nbsp;</td>")
	  Response.Write("<td class='rText'>" & objRS10("PRPD" & recnum) & "</td>")
	  Response.Write("<td class='rText2'>&nbsp;</td>")
	  Response.Write("<td align='right' class='rText'>" & calcDateRS10(("SRAUDT"), recnum) & "</td>")
	  Response.Write("<td class='rText2'>&nbsp;</td>")
	End if
	End if
  Next

	else
		Response.Write("<td width='80' class='rText' colspan='7'>No Sales </td>")
    end if

  %>


</table>


