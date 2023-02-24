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
<table width="650" border="0" cellpadding="0" cellspacing="0">
  <tr valign="top">
    <td width="40" class="SHeader">&nbsp;</td>
    <td width="4" class="SHeader"></td>
    <td width="90" class="SHeader"></td>
    <td width="4" class="SHeader"></td>
    <td width="40" class="SHeader"></td>
    <td width="4" class="SHeader"></td>
    <td width="40" class="SHeader"></td>
    <td width="4" class="SHeader"></td>
    <td width="40" class="SHeader"></td>
    <td width="4" class="SHeader"></td>
    <td width="40" class="SHeader"></td>
    <td width="4" class="SHeader"></td>




  </tr>
  <tr valign="top">
    <td colspan="15"></td>
  </tr>
  <tr valign="top">


    <td class="STitle" align='right'>Code</td>
    <td class="STitle"></td>
    <td class="STitle" align='center'>Code Desc</td>
    <td class="STitle"></td>
    <td class="STitle" align='right'>Curr-Prin</td>
    <td class="STitle"></td>
    <td class="STitle" align='right'>Curr-Int</td>
    <td class="STitle"></td>
    <td class="STitle" align='right'>Curr-Due</td>
    <td class="STitle"></td>
    <td class="STitle" align='right'>Bal-Rem</td>
    <td class="STitle"></td>


  </tr>
   <%
   recnum = 1
    For recnum = 1 To 5  ' used to be 10 and there were errors from it. Checked the MDB and these only go to 5.

    	if objRSSPT("TSSAC1" & recnum) <> "" Then
    	if objRSSPT("TSSAC1" & recnum) > 0 then
    	result = recnum MOD 2
    	If result = 0 then
	      Response.Write("<tr valign='top'>")
	  	  Response.Write("<td class='rText2'>"& objRSSPT("TSSAC1" & recnum) & "</td>")
	  	  Response.Write("<td class='rText'></td>")
	  	  Response.Write("<td class='rText2'>"& objRSSPT("TSSAD1" & recnum) & "</td>")
	  	  Response.Write("<td class='rText'></td>")
	  	  Response.Write("<td align='right' class='rText2'>" & Formatnumber(objRSSPT("TSSAP1" & recnum), 2) & "</td>")
	  	  Response.Write("<td class='rText'></td>")
	  	  Response.Write("<td align='right' class='rText2'>" &  Formatnumber(objRSSPT("TSSAI1" & recnum), 2) & "</td>")
	  	  Response.Write("<td class='rText'></td>")
	  	  Response.Write("<td align='right' class='rText2'>" &  Formatnumber(objRSSPT("TSSAP1" & recnum) + objRSSPT("TSSAI1" & recnum), 2) & "</td>")
	  	  Response.Write("<td class='rText'></td>")
	  	  Response.Write("<td align='right' class='rText2'>" &  Formatnumber(objRSSPT("TSSAR1" & recnum), 2) & "</td>")
	  	  Response.Write("<td class='rText'></td>")
		  Response.Write("<td class='rText'></td>")
		  Response.Write("<td class='rText'></td>")
	  	  Response.Write("</tr>")
	  	Else
	      Response.Write("<tr valign='top'>")
	  	  Response.Write("<td class='rText'>"& objRSSPT("TSSAC1" & recnum) & "</td>")
	  	  Response.Write("<td class='rText'></td>")
	  	  Response.Write("<td class='rText'>"& objRSSPT("TSSAD1" & recnum) & "</td>")
	  	  Response.Write("<td class='rText'></td>")
	  	  Response.Write("<td align='right' class='rText'>" & Formatnumber(objRSSPT("TSSAP1" & recnum), 2) & "</td>")
	  	  Response.Write("<td class='rText'></td>")
	  	  Response.Write("<td align='right' class='rText'>" &  Formatnumber(objRSSPT("TSSAI1" & recnum), 2) & "</td>")
	  	  Response.Write("<td class='rText'></td>")
	  	  Response.Write("<td align='right' class='rText'>" &  Formatnumber(objRSSPT("TSSAP1" & recnum) + objRSSPT("TSSAI1" & recnum), 2) & "</td>")
	  	  Response.Write("<td class='rText'></td>")
	  	  Response.Write("<td align='right' class='rText'>" &  Formatnumber(objRSSPT("TSSAR1" & recnum), 2) & "</td>")
	  	  Response.Write("<td class='rText'></td>")
		  Response.Write("<td class='rText'></td>")
		  Response.Write("<td class='rText'></td>")
	  	  Response.Write("</tr>")
	  	End if
  	End if
  	Else
	If recnum < 2 then
		Response.Write("<td class='rText2'>No Values</td>")
			end if
    	Exit For
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