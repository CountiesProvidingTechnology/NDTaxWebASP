

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
		<td class="rText"><%= objRSPYGEN("TXPRCL") %></td>
		<td class="STitle">Name</td>
		<td class="rText"><%= objRSPYGEN("TPNAME") %></td>
	</tr>
	<tr>
		<td class="STitle"></td>
		<td class="rText"></td>
		<td class="STitle"></td>
		<td class="rText"></td>
	</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0">
<td width="200" ></td>
<td width="200" ></td>
<td width="200" ></td>
<td width="250" align="right" class="tText2">This data is subject to change</td>
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
    <td width="10" class="SHeader"></td>
    <td width="4" class="SHeader"></td>
	<td width="10" class="SHeader"></td>
	<td width="4" class="SHeader"></td>
    <td width="10" class="SHeader"></td>
    <td width="4" class="SHeader"></td>
    <td width="20" class="SHeader"></td>
  </tr>
  <tr valign="top">
    <td colspan="25"></td>
  </tr>
  <tr valign="top">
    <td class="STitle">Code</td>
    <td class="STitle"></td>
    <td class="STitle">Init Amount</td>
    <td class="STitle"></td>
    <td class="STitle">Rate</td>
    <td class="STitle"></td>
    <td class="STitle">Start Year</td>
    <td class="STitle"></td>
    <td align="right" class="STitle">Number of Years</td>
    <td class="STitle"></td>
    <td class="STitle">Extra</td>
    <td class="STitle"></td>
    <td class="STitle">Curr Prin</td>
    <td class="STitle"></td>
    <td class="STitle">Curr Int</td>
    <td class="STitle"></td>
    <td class="STitle">Bal Rem</td>
    <td class="STitle"></td>
    <td align='right' class="STitle">DP</td>
    <td class="STitle"></td>
    <td align='right' class="STitle">DI</td>
    <td class="STitle"></td>
    <td align='right' class="STitle">IN</td>
    <td class="STitle"></td>
    <td align='right' class="STitle">Type</td>
    <td class="STitle"></td>
  </tr>
   <%
   recnum = 1
    For recnum = 1 To 10
    'Response.Write(" the value of recnum : " & recnum )
		if cid = 30 then
    	if strTwn <> 65 Then

  		Response.Write("<td class='rText2'>No Values</td>")

    		Exit For
    	end if
		end if
    	if objRS11("INIT0" & recnum) = 0 Then
  		If recnum < 2 then
  		Response.Write("<td class='rText2'>No Values</td>")
  		end if
    		Exit For
    	else


    	result = recnum MOD 2
    	If result = 0 then
        Response.Write("<tr valign='top'>")
  	  Response.Write("<td class='rText2'>"& objRS11("SPCD0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>"& Formatnumber(objRS11("INIT0" & recnum), 2) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" &  Formatnumber(objRS11("RATE0" & recnum), 2) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  if len(objRS11("YEAR0" & recnum)) = 1 then
  	  Response.Write("<td align='right' class='rText2'>0" & objRS11("YEAR0" & recnum) & "</td>")
  	  else
  	  Response.Write("<td align='right' class='rText2'>" & objRS11("YEAR0" & recnum) & "</td>")
  	  end if
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & objRS11("NUMB0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & objRS11("XTRA0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & Formatnumber(objRS11("CPRN0" & recnum), 2) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & Formatnumber( objRS11("CINT0" & recnum), 2) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & Formatnumber(objRS11("BREM0" & recnum), 2) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" &  objRS11("DPFL0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & objRS11("DIFL0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText2'>" & objRS11("INFL0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td class='rText2'>" & objRS11("TYPE0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("</tr>")
  	Else
        Response.Write("<tr valign='top'>")
  	  Response.Write("<td class='rText'>"& objRS11("SPCD0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>"& Formatnumber(objRS11("INIT0" & recnum), 2) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" &  Formatnumber(objRS11("RATE0" & recnum), 2) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  if len(objRS11("YEAR0" & recnum)) = 1 then
  	  Response.Write("<td align='right' class='rText'>0"& objRS11("YEAR0" & recnum) & "</td>")
  	  else
  	  Response.Write("<td align='right' class='rText'>"& objRS11("YEAR0" & recnum) & "</td>")
  	  end if
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>"& objRS11("NUMB0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>"& objRS11("XTRA0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" & Formatnumber(objRS11("CPRN0" & recnum), 2) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" & Formatnumber( objRS11("CINT0" & recnum), 2) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" & Formatnumber(objRS11("BREM0" & recnum), 2) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" &  objRS11("DPFL0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" & objRS11("DIFL0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
  	  Response.Write("<td align='right' class='rText'>" & objRS11("INFL0" & recnum) & "</td>")
  	  Response.Write("<td class='rText'></td>")
	  Response.Write("<td align='right' class='rText'>" & objRS11("TYPE0" & recnum) & "</td>")
	  Response.Write("<td class='rText'></td>")

  	  Response.Write("</tr>")
  	End if
 	End if
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


