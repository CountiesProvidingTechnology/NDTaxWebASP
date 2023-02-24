

<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="80" class="SHeader"></td>
		<td width="200" class="SHeader">Special Assessment Amounts</td>
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
		<td class="rText"><%= objRS("TPNAME") %></td>
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
  </tr>
  <tr valign="top">
    <td colspan="15"></td>
  </tr>
  <tr valign="top">
    <td class="STitle">Type of Project</td>
    <td class="STitle"></td>
    <td class="STitle">Project Number</td>
    <td class="STitle"></td>
    <td class="STitle">Original Assessment</td>
    <td class="STitle"></td>
    <td class="STitle">Start Year</td>
    <td class="STitle"></td>
    <td align="right" class="STitle">Number of Years</td>
    <td class="STitle"></td>
    <td class="STitle">Interest Rate</td>
    <td class="STitle"></td>
    <td class="STitle">Remaining Balance</td>
    <td class="STitle"></td>
    <td class="STitle">Estimated Annual Principle</td>
    <td class="STitle"></td>
    <td class="STitle">Estimated Annual Interest</td>
    <td class="STitle"></td>
    <td align='right' class="STitle">Estimated Annual Installment</td>
    <td class="STitle"></td>
  </tr>
   <%
   recnum = 1
    For recnum = 1 To 10
'   Response.Write(" the value of recnum : " & recnum )
'Response.Write(" the value of strTwn : " & strTwnrecnum )
'    	if strTwn <> 65 Then
'
'  		Response.Write("<td class='rText2'>No Values</td>")
'
'    		Exit For
'    	end if
    	result = recnum MOD 2
	If result = 0 then

       Response.Write("<tr valign='top'>")
  	  Response.Write("<td class='rText2'>"& objRSSP("TSSAC1" & recnum) & "</td>")
' 	  Response.Write("<td align='right' class='rText2'>"& objRS11("SPCD0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" &  Formatnumber(objRSSP("ASAMT" & recnum), 2) & "</td>")
'  	  Response.Write("<td align='right' class='rText2'>" &  objRS11("XTRA0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td class='rText2'>" & objRSSP("ASRAT" & recnum) & "</td>")
'  	  Response.Write("<td align='right' class='rText2'>" & objRS11("INIT0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td class='rText2'>" & objRSSP("ASSTR" & recnum) & "</td>")
'  	  Response.Write("<td align='right' class='rText2'>" & objRS11("YEAR0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & objRSSP("ASTRS" & recnum) & "</td>")
'  	  Response.Write("<td align='right' class='rText2'>" & objRS11("NUMB0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & objRSSP("ASXTR" & recnum) & objRSSP("ASXDS" & recnum) & "</td>")
'  	  Response.Write("<td align='right' class='rText2'>" & objRS11("RATE0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & Formatnumber( objRSSP("TSSAP" & recnum), 2) & "</td>")
'  	  Response.Write("<td align='right' class='rText2'>" & objRS11("BREM0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & objRSSP("TSSAI" & recnum) & "</td>")
'  	  Response.Write("<td align='right' class='rText2'>" & objRS11("CPRN0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" &  objRSSP("TSSAR" & recnum) & "</td>")
'  	  Response.Write("<td align='right' class='rText2'>" &  objRS11("CINT0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & objRSSP("ASDPF" & recnum) & "</td>")
'		esttotal =	objRS11("CINT0" & recnum) + objRS11("CPRN0" & recnum)
'  	  Response.Write("<td align='right' class='rText2'>" & esttotal & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & objRSSP("ASDIF" & recnum) & "</td>")
'  	  Response.Write("<td class='rText'></td>")
'  	  Response.Write("<td align='right' class='rText2'>" & objRSSP("ASINF" & recnum) & "</td>")
'  	  Response.Write("<td class='rText'></td>")
'  	  Response.Write("<td class='rText2'>" & objRSSP("TSSAY" & recnum) & "</td>")
'  	  Response.Write("<td class='rText'></td>")
 	  Response.Write("</tr>")
  	Else
        Response.Write("<tr valign='top'>")
  	  Response.Write("<td class='rText'>"& objRSSP("TSSAC1" & recnum) & "</td>")
'  	  Response.Write("<td align='right' class='rText'>"& objRS11("SPCD0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" &  Formatnumber(objRSSP("ASAMT" & recnum), 2) & "</td>")
'  	  Response.Write("<td align='right' class='rText'>" &  objRS11("XTRA0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td class='rText'>"& objRSSP("ASRAT" & recnum) & "</td>")
'  	  Response.Write("<td align='right' class='rText'>"& objRS11("INIT0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td class='rText'>"& objRSSP("ASSTR" & recnum) & "</td>")
'  	  Response.Write("<td align='right' class='rText'>"& objRS11("YEAR0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>"& objRSSP("ASTRS" & recnum) & "</td>")
'  	  Response.Write("<td align='right' class='rText'>"& objRS11("NUMB0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" & objRSSP("ASXTR" & recnum) & objRSSP("ASXDS" & recnum) & "</td>")
'  	  Response.Write("<td align='right' class='rText'>" & objRS11("RATE0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" & Formatnumber( objRSSP("TSSAP" & recnum), 2) & "</td>")
'  	  Response.Write("<td align='right' class='rText'>" & objRS11("BREM0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" & objRSSP("TSSAI" & recnum) & "</td>")
'  	  Response.Write("<td align='right' class='rText'>" & objRS11("CPRN0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" &  objRSSP("TSSAR" & recnum) & "</td>")
'  	  Response.Write("<td align='right' class='rText'>" &  objRS11("CINT0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
'		esttotal =	objRS11("CINT0" & recnum) + objRS11("CPRN0" & recnum)
'		response.write(" the value of the total is : " & esttotal )
'  	  Response.Write("<td align='right' class='rText'>" & esttotal & "</td>")
  	  Response.Write("<td class='rText'></td>")
'  	  Response.Write("<td align='right' class='rText'>" & objRSSP("ASDIF" & recnum) & "</td>")
'  	  Response.Write("<td class='rText'></td>")
'	  Response.Write("<td align='right' class='rText'>" & objRSSP("ASINF" & recnum) & "</td>")
'	  Response.Write("<td class='rText'></td>")
'	  Response.Write("<td class='rText'>" & objRSSP("TSSAY" & recnum) & "</td>")
'	  Response.Write("<td class='rText'></td>")
  	  Response.Write("</tr>")
  	End if
'  	End if
    Next
  %>

  	<tr>
  		<td ></td>
  		<td >&nbsp;</td>
  		<td ></td>
  		<td ></td>
	</tr>
	</tr>
	  	<td ></td>
	  	<td >&nbsp;</td>
	  	<td ></td>
	  	<td ></td>
	</tr>


</table>


