

<table border="0" cellpadding="0" cellspacing="0">
	<tr colspan="6" valign="top">
		<td width="280" class="SHeader">Appraisal Summary</td>
		<td width="80" class="SHeader"></td>
		<td width="70" class="SHeader"></td>
		<td width="150" class="SHeader"></td>
		<td width="70" class="SHeader"></td>
		<td width="250" class="SHeader"></td>
	</tr>
	<tr valign="top">
		<td colspan="6" height="1" bgcolor="#000000"></td>
	</tr>
	<tr>
		The <%= objRSPYGEN("TXYEAR") %> assessment reflects the property value as of January 2nd, <%= objRSPYGEN("TXYEAR") %> using sales that
		occurred between October <%= objRSPYGEN("TXYEAR") - 2 %> and September <%= objRSPYGEN("TXYEAR")- 1 %>. Buildings built prior to January 2nd, <%= objRSPYGEN("TXYEAR") %> or buildings which were partially
		complete as of January 2nd, <%= objRSPYGEN("TXYEAR") %> are included here. Any buildings built after January 2nd, <%= objRSPYGEN("TXYEAR") %>
		will be included on the January 2nd, <%= (objRSPYGEN("TXYEAR")+1) %> assessment.

	</tr>
	<tr valign="top">
		<td width="100" class="rText">Parcel Number </td>
		<td width="180" class="rText"><%= objRS("TXPRCL") %></td>
		<td width="1" class="rText"></td>
		<td width="100" class="rText"><%= objRS("TXCITYD") %></td>
		<td width="1" class="rText"></td>
		<td width="250" class="rText"><%= objRS("TXSCHLD") %>&nbsp;<%= objRS("TXSCHL") %></td>
</table>
<table width="600" border="0" cellpadding="0" cellspacing="0">
	</tr>
	<tr>
		<td width="280" class="STitle">Primary Taxpayer</td>
		<td width="1" class="STitle"></td>
		<td width="280" class="STitle">Legal Description</td>
		<td width="1" class="STitle"></td>
		<td width="280" class="STitle"></td>
	</tr>
	<tr>
		<td width="280" valign="top" class="rText"><%= objRSPYGEN("TXTNAM") %><br>
						  <%= objRSPYGEN("TXTAD1") %><br>
						  <%= objRSPYGEN("TXTAD2") %><br>
						  <%= objRSPYGEN("TXTAD3") %><br>
						  <%= objRSPYGEN("TXTAD4") %><br>
		</td>
		<td width="1" class="rText">&nbsp;&nbsp; </td>
		<td width="280" class="rText">
		<% 	If objRSPYGEN("TXSECT") > 0 Then %>
			Sect - <%= objRSPYGEN("TXSECT") %>
		<% end if %>
		<% 	If objRSPYGEN("TXTOWN") > 0 Then %>
			Twp - <%= objRSPYGEN("TXTOWN") %>
		<% end if %>
		<% 	If objRSPYGEN("TXRANG") > 0 Then %>
			Range - <%= objRSPYGEN("TXRANG") %>
		<% end if %>
		<% 	If objRSPYGEN("TXLOT") > 0 Then %>
			Lot - <%= objRSPYGEN("TXLOT") %>
		<% end if %>
		<% 	If objRSPYGEN("TXBLOK") > 0 Then %>
			Block - <%= objRSPYGEN("TXBLOK") %>
		<% end if %><br>
						  <%= objRSPYGEN("TXPLATD") %><br>
						  <%= objRSPYGEN("TXDSC1") %><br>
						  <%= objRSPYGEN("TXDSC2") %><br>
						  <%= objRSPYGEN("TXDSC3") %><br>
						  <%= objRSPYGEN("TXDSC4") %><br>
						  <%= objRSPYGEN("TXDSC5") %><br>
						  <%= objRSPYGEN("TXDSC6") %><br>
						  <%= objRSPYGEN("TXDSC7") %><br>
						  <%= objRSPYGEN("TXDSC8") %><br>
		</td>
		<td width="1" class="rText">&nbsp;&nbsp; </td>
		<td width="280" class="rText"></td>

<%  ' Start the loop for the table 11 legal descriptions here   ! ! ! ! !  !


if objRS11.EOF Then

else
              While objRS11("TXREC") < 99
		Response.Write("<tr>")
		Response.Write("<td width='280' class='rText'></td>")
		Response.Write("<td class='rText'></td>")
		Response.Write("<td width='280' class='rText'>" & objRS11("TXDESC") & "</td>")
		Response.Write("</tr>")

	objRS11.MoveNext

Wend



End If
%>

		</td>
	</tr>

		<% If objRSPYGEN("TXALTR") > 0 Then %>
		<td width="280" class="STitle">Alternate Mailing Address</td>
		<td></td>
		<td width="280" class="STitle"></td>
	  </tr>
		<% Else %>
				<td width="280" class="STitle"></td>
				<td></td>
				<td width="280" class="STitle"></td>
				<td></td>
				<td width="280" class="STitle"></td>
		<% End If %>
	<tr>
<% If objRSPYGEN("TXALTR") > 0 Then %>
				<%
						Response.Write("<td width='240' class='rText'>")
						Response.Write(objRSPYGEN("TXANAM") & "<br>")
						Response.Write(objRSPYGEN("TXAAD1") & "<br>")
						Response.Write(objRSPYGEN("TXAAD2") & "<br>")
						Response.Write(objRSPYGEN("TXAAD3") & "<br>")
						Response.Write(objRSPYGEN("TXAAD4") & "</td>")
				%>
		<td width="1"> </td>
		<td width="280" class="rText">
		<%
%>
		</td>

<% Else




End If %>
		<%'=' objRSPYGEN("TXPADR1") %>&nbsp;&nbsp;<%'=' objRSPYGEN("TXPZIP1") %>
	</tr>

	<tr>
		<td class="STitle">Property Classification</td>
		<td class="STitle"></td>
		<td class="STitle">Property Address</td>
		<td class="STitle"></td>
		<td class="STitle">Lake #</td>
	</tr>
	<tr>
<% If Not objRSCAMA.EOF Then




%><% 'If there are no records Then do not show the information below' %>

		<td width="280" class="rText"><%= objRSCAMA("ASMT1") %><br>
						  <%= objRSCAMA("ASMT2") %><br>
						  <%= objRSCAMA("ASMT3") %><br>
		</td>
		<td class="rText"></td>
		<td width="280" valign="top" class="rText">
		<%
		If objRSPYGEN("TXPADR1") <> "" Then

		Response.Write(objRSPYGEN("TXPADR1"))
		End If
		%>&nbsp;&nbsp;
		<%
		If objRSPYGEN("TXPADR1") <> "" Then
			If objRSPYGEN("TXPZIP1") = "00000"  Then
			Response.Write(" ")
			Else
				If objRSPYGEN("TXPZIP1") = "000000000" Then
				Response.Write(" ")
				Else
				'Response.Write("here at last")
				Response.Write(objRSPYGEN("TXPZIP1"))
				End if
			End If
		End If
		%>
		<%'=' objRSPYGEN("TXPADR1") %>&nbsp;&nbsp;<%'=' objRSPYGEN("TXPZIP1") %>


		</td>

<td class="rText"></td>
		<td width="280" valign="top" class="rText">
		<%

		If not objRS11.EOF  then
		If objRS11("TXREC") = 99  Then     '   the reach into the Table 11 looking at the 99 record for a lake description

		Response.Write(objRS11("TXDESC"))
		else

		End If
		Else

		End If
		%>&nbsp;&nbsp;
		<%

		%>
		<%'=' objRSPYGEN("TXPADR1") %>&nbsp;&nbsp;<%'=' objRSPYGEN("TXPZIP1") %>


		</td>

	<tr valign="top">
		<td ></td>
<%


'  ****  this line will handle the new CAMA PDF button to show/Print the same information as is seen on the web .  3-3-10 LEM'
Response.Write("<td width='600' align='right' colspan='2'></td>")
response.Write("<td width='600' align='right' colspan='2'><A HREF='http://cpuimei.com:41080/iText/MNPAS2010.jsp?cid=" & varcid & "&pid=" & strPID  & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntAppraslInfo.gif'></A></td>")


%>
	</tr>
<% Else %>
		<td width="280" class="rText">No records for this parcel</td>
		<td class="rText"></td>
		<td width="280" class="rText"></td>
		<td class="rText"></td>
		<td width="280" class="rText"></td>

<% End If





%>
	</tr>
</table>

<% If Not objRSCAMA.EOF Then %>

<table width="600" border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td colspan="4" height="1" bgcolor="#000000"></td>
	</tr>
<tr>
		<td class="rText">Estimated Market Value</td>
		<td class="rText"></td>
		<td class="rText" align="right"><%= FormatNumber(objRSCAMA("TotMkt"), 0) %></td>
		<td width="10" class="rText"></td>
		<td width="240" class="STitle">Primary House Summary</td>
		<td class="STitle"></td>
	</tr>
	<tr>
		<td  class="rText">Exempt Wetlands/Native Prairie</td>
		<td class="rText"></td>
		<td class="rText" align="right" ><%= calcZero_BlankCAMA("TotWet", 0) %></td>
		<td width="10" class="rText"></td>
		<td class="rText">Condition</td>
		<td  class="rText" align="right" ><%= objRSCAMA("Housecond") %></td>
	</tr>
	<tr>
		<td  class="rText">Green Acres Value Def</td>
		<td class="rText"></td>
		<td class="rText" align="right" ><%= calcZero_BlankCAMA("TotGA", 0) %></td>
		<td width="10" class="rText"></td>
		<td class="rText">Type</td>
		<td  class="rText" align="right" ><%= objRSCAMA("HouseType") %></td>
	</tr>
	<tr>
		<td  class="rText">Rural Pres Value Deferred</td>
		<td class="rText"></td>
		<td class="rText" align="right" ><%= calcZero_BlankCAMA("TotRrlp", 0) %></td>
		<td width="10" class="rText"></td>
		<td class="rText"># of Units</td>
		<td  class="rText" align="right" ></td>
	</tr>
	<tr>
		<td  class="rText">Plat Deferment</td>
		<td class="rText"></td>
		<td class="rText" align="right" ><%= calcZero_BlankCAMA("TotPlat", 0) %></td>
		<td width="10" class="rText"></td>
		<td class="rText">Total Sq Ft</td>
		<td  class="rText" align="right" ><%= objRSCAMA("Sqft") %></td>
	</tr>
	<tr>
		<td  class="rText">JOBZ Amount Exempted</td>
		<td class="rText"></td>
		<td class="rText" align="right" ><%= calcZero_BlankCAMA("TotJOBZ", 0) %></td>
		<td width="10" class="rText"></td>
		<td class="rText">Year Built</td>
		<td  class="rText" align="right" ><%= calcyearCAMA("YrBlt") %></td>
	</tr>
	<tr>
		<td  class="rText">This Old House Exclusion</td>
		<td class="rText"></td>
		<td class="rText" align="right" ><%= calcZero_BlankCAMA("TotHous", 0) %></td>
		<td width="10" class="rText"></td>
		<td class="rText">Year Remodel</td>
		<td  class="rText" align="right" ><%= calcyearCAMA("YrRemodel") %></td>
	</tr>
	<tr>
		<td  class="rText">Dis Vets Mkt Value Excl</td>
		<td class="rText"></td>
		<td class="rText" align="right" ><%= calcZero_BlankCAMA("TotVets", 0) %></td>
		<td width="10" class="rText"></td>
		<td class="rText">Air Cond</td>
		<td  class="rText" align="right" ><%= objRSCAMA("AC") %></td>
	</tr>
	<tr>
		<td  class="rText">Homestead Mkt Value Excl</td>
		<td class="rText"></td>
		<td class="rText" align="right" ><%= calcZero_BlankCAMA("TotExcl", 0) %></td>
		<td width="10" class="rText"></td>
		<td class="rText"></td>
		<td  class="rText" align="right" ></td>
	</tr>
	<tr>
		<td  class="rText">Taxable Market Value</td>
		<td class="rText"></td>
		<td class="rText" align="right" ><%= FormatNumber(objRSCAMA("TotTMV"), 0) %></td>
		<td width="10" class="rText"></td>
		<td class="rText"></td>
		<td  class="rText" align="right" ></td>
	</tr>
	<tr>
		<td  class="rText">New Improvements incl. in Est Mkt</td>
		<td class="rText"></td>
		<td class="rText" align="right" ><%= calcZero_BlankCAMA("TotImpv", 0) %></td>
		<td width="10" class="rText"></td>
		<td class="rText"></td>
		<td  class="rText" align="right" ></td>
	</tr>
	<tr>
		<td  class="rText">Referendum Market Val</td>
		<td class="rText"></td>
		<td class="rText" align="right" ><%= calcZero_BlankCAMA("ValOpn1", 0) %></td>
		<td width="10" class="rText"></td>
		<td class="rText"></td>
		<td  class="rText" align="right" ></td>
	</tr>
</table>


<table width="600" border="0" cellpadding="0" cellspacing="0">

	<tr>
<%
		strData=objRS("TXPRCL")
	'response.write("the value of ParcelNo : " & strData )
		strtrimdata = Trim(strData)
	intLength = Len(strtrimdata)
	If varcid = 21 or varcid = 26 or varcid=41 or varcid=45 or varcid=53 or varcid=61 or varcid=67 or varcid=75 or varcid=76  Then         'Creates  formatted parcel XX-XXXX-XXX  (2-4-3)
	strleftChars = Left(strtrimdata, 2)
	strmidChars = Mid(strtrimdata, 4, 4)
	strrightChars = Right(strtrimdata, 3)
	strParcelchar = strleftChars + strmidChars + strrightChars + "0101.JPG"
	strParcelchar = "\\206.145.187.205\sketches\" + varcid + "\" + strParcelchar
	strParcelchar1 = strleftChars + strmidChars + strrightChars + "0101-1.JPG"
	strParcelchar1 = "\\206.145.187.205\sketches\" + varcid + "\" + strParcelchar1
	strParcelchar2 = strleftChars + strmidChars + strrightChars + "0201.JPG"
	strParcelchar2 = "\\206.145.187.205\sketches\" + varcid + "\" + strParcelchar2
	strParcelchar3 = strleftChars + strmidChars + strrightChars + "0201-1.JPG"
	strParcelchar3 = "\\206.145.187.205\sketches\" + varcid + "\" + strParcelchar3
	end if

	'Response.Write(" the value of the parcel var1: " & strParcelchar1 )


Dim objFSO, objFile, objFSO1, objFSO2, objFSO3
set objFSO = CreateObject("Scripting.FileSystemObject")
set objFSO1 = CreateObject("Scripting.FileSystemObject")
set objFSO2 = CreateObject("Scripting.FileSystemObject")
set objFSO3 = CreateObject("Scripting.FileSystemObject")





%>




	</tr>
</table>
<table width="600" border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td colspan="11" height="1" bgcolor="#000000"></td>
	</tr>
  <tr valign="top">
    <td colspan="11"></td>
  </tr>
  <tr valign="top">
    <td class="STitle">Effective Yr</td>
    <td width="1"></td>
    <td class="STitle">Item</td>
    <td class="rText"></td>
    <td class="STitle">Type</td>
    <td class="rText"></td>
    <td class="STitle">Quantity/SF</td>
    <td class="rText"></td>
    <td class="STitle" align="center">CER</td>
  </tr>
<%
If Not objRSCAMASum.EOF Then
tmpRecNum = objRSCAMASum("RecNum")
tmpSecNum = objRSCAMASum("SecNum")
End If
While Not objRSCAMASum.EOF
	if objRSCAMASum("Item") <> ""     Then
		Response.Write("<tr valign='top'>")
		  Response.Write("<td align='center' width='80' class='rText'>" & calcZeroCAMASum("Yreff") & "</td>")

		  Response.Write("<td width='1'></td>")
		  Response.Write("<td width='90' class='rText'>" & objRSCAMASum("Item") & "</td>")
		  Response.Write("<td width='1'></td>")
		  Response.Write("<td  width='90' class='rText'>" & objRSCAMASum("Type") & "</td>")
		  Response.Write("<td width='1'></td>")
		  Response.Write("<td align='right' width='50' class='rText'>" & calcZeroCAMASum2("Sqft", 2) & "</td>")
		  Response.Write("<td width='1'></td>")
		  Response.Write("<td width='90' align='center' class='rText'>" & calcZero_BlankCAMASum("CER", 2) & "</td>")
		Response.Write("</tr>")
  End If
	objRSCAMASum.MoveNext

Wend

objRSCAMASum.Close
Set objRSCAMASum = Nothing
%>
</table>
<table width="600" border="0" cellpadding="0" cellspacing="0">
  <tr valign="top">
    <td class="STitle" width='60'>Totals</td>
    <td class="STitle" width='10'></td>
    <td class="STitle" width='10'></td>
    <td class="STitle" width='10'></td>
    <td class="STitle" width='60'></td>
    <td class="STitle" width='10'></td>
    <td class="STitle" width='60'></td>
	<td class="STitle" width='10'></td>
    <td class="STitle" width='60'></td>
	<td class="STitle" width='10'></td>
    <td class="STitle" width='60'></td>
    <td class="STitle" width='100'></td>
  </tr>
  <tr valign="top">
    <td class="rText">Land</td>
    <td width='1'></td>
    <td class="rText"><%= FormatNumber(objRSCAMA("LandVal"), 0) %></td>
    <td width='1'></td>
    <td class="rText">Building</td>
    <td width='1'></td>
    <td class="rText"><%= FormatNumber(objRSCAMA("BldVal"), 0) %></td>
    <td width='1'></td>
    <td class="rText">Total</td>
    <td width='1'></td>
    <td class="rText"><%= FormatNumber(objRSCAMA("TotVal"), 0) %></td>
  </tr>
</table>
		<% End If %>
<% if varcid = 21 then %>
<img src=http:<%= strParcelchar %> onerror="this.onerror=null;this.src='http://206.145.187.205/sketches/missing.gif';">





<img  src=http:<%= strParcelchar1 %> onerror="this.onerror=null;this.src='http://206.145.187.205/sketches/missing.gif';">




<img src=http:<%= strParcelchar2 %> onerror="this.onerror=null;this.src='http://206.145.187.205/sketches/missing.gif';">



<img  src=http:<%= strParcelchar3 %> onerror="this.onerror=null;this.src='http://206.145.187.205/sketches/missing.gif';">

<% end If %>
<% if varcid = 48 then %>
<img src=http:<%= strParcelchar %> onerror="this.onerror=null;this.src='http://206.145.187.205/sketches/missing.gif';">





<img  src=http:<%= strParcelchar1 %> onerror="this.onerror=null;this.src='http://206.145.187.205/sketches/missing.gif';">




<img src=http:<%= strParcelchar2 %> onerror="this.onerror=null;this.src='http://206.145.187.205/sketches/missing.gif';">



<img  src=http:<%= strParcelchar3 %> onerror="this.onerror=null;this.src='http://206.145.187.205/sketches/missing.gif';">

<% end If %>
<% if varcid = 61 then %>
<img src=http:<%= strParcelchar %> onerror="this.onerror=null;this.src='http://206.145.187.205/sketches/missing.gif';">





<img  src=http:<%= strParcelchar1 %> onerror="this.onerror=null;this.src='http://206.145.187.205/sketches/missing.gif';">




<img src=http:<%= strParcelchar2 %> onerror="this.onerror=null;this.src='http://206.145.187.205/sketches/missing.gif';">



<img  src=http:<%= strParcelchar3 %> onerror="this.onerror=null;this.src='http://206.145.187.205/sketches/missing.gif';">

<% end If %>


