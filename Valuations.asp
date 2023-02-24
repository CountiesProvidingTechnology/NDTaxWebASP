<%
Session("strPID") = objRS95("Parcel")
sesParcel = Session("strPID")
'Response.Write("sesRecNumV#1= ") & sesRecNum

counter = request.form("first")
'counter = counter + 1'

counter = counter + 1
Session("counter") = counter
sesRecNum = Session("counter")
sesTotalRec = intTotalRec
'Response.write("sesRecNum= ") & sesRecNum
'Response.Write("intI5= ") & intI5'
%>

	<% If sesRecNum > intTotalRec Then
	sesRecNum = 1
	End If %>


<table border="0" cellpadding="0" cellspacing="0">
  <tr valign="top">
    <td width="100" class="SHeader" "nowrap" ><font class="oText"><%= objRS95("MH") %></font></td>
    <td width="60" class="SHeader" "nowrap" ><font class="oText">&nbsp;</font>&nbsp;MP&nbsp;#</td>
    <td width="80" class="rText"><%= objRS95("MULTI") %>&nbsp;&nbsp;</td>
    <td width="40"  class="SHeader">&nbsp;Name&nbsp;</td>
    <td width="200" class="rText"><%= objRS95("Taxpayer") %></td>
    <% intTotalRec = intI5 %>


    <% If intI5 <> 1 Then %>
    <td width="100" "nowrap" class="rText2">Record&nbsp;<%= (Session("counter")-1) %> &nbsp;of&nbsp;<%= intTotalRec %></td>
    <% Else %>
    <td width="100" "nowrap" class="rText2">&nbsp;&nbsp;&nbsp; </td>
    <% End If %>

  </tr>
  </table>
<table border="0" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
  <tr valign="top">
    <td colspan="6" height="1" bgcolor="#000000"></td>
    <td width="22" bgcolor="#000000"></td>
    <td width="45" bgcolor="#000000"></td>
    <td width="35" bgcolor="#000000"></td>
    <td width="40" bgcolor="#000000"></td>
  </tr>
  <tr valign="top">
    <td width="80" class="SHeader">Year</td>
    <% If objRS95.EOF Then %>
		<td colspan="2" align="center" class="STitle">&nbsp;</td>
    <% Else %>
    	<td colspan="2" align="center" class="STitle"><%= objRS95("YEAR") %>&nbsp;</td>
    <% End If %>
    <td width="20" align="right" class="STitle"></td>
    <% If objRS94.EOF Then %>
    	<td colspan="2" align="center" class="STitle">&nbsp;</td>
    <% Else %>
    	<td colspan="2" align="center" class="STitle"><%= objRS94("YEAR") %>&nbsp;</td>
    <% End If %>
    <td width="20" class="STitle" align="right"></td>
    <% If objRS93.EOF Then %>
    	<td colspan="2" align="center" class="STitle">&nbsp;</td>
    <% Else %>
    	<td colspan="2" align="center" class="STitle"><%= objRS93("YEAR") %>&nbsp;</td>
    <% End If %>
    <td class="SHeader"></td>
  </tr>
  <tr valign="top">
    <td colspan="6" height="1" bgcolor="#000000"></td>
    <td bgcolor="#000000"></td>
    <td bgcolor="#000000"></td>
    <td bgcolor="#000000"></td>
    <td bgcolor="#000000"></td>
  </tr>
  <tr valign="top">
    <td class="STitle">ASMT&nbsp;Code/Desc</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td colspan="2" class="rText">&nbsp;<%= objRS95("ASMT") %>&nbsp;<%= objRS95("ASMTDESC") %></td>
    <% End If %>
    	<td align="right" class="STitle"></td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td colspan="2" class="rText">&nbsp;<%= objRS94("ASMT") %>&nbsp;<%= objRS94("ASMTDESC") %></td>
    <% End If %>
    	<td class="STitle" align="right"></td>
    <% If  objRS93.EOF Then %>
    	<td colspan="2" class="rText">&nbsp;&nbsp;</td>
    <% Else %>
    	<td colspan="2" class="rText">&nbsp;<%= objRS93("ASMT") %>&nbsp;<%= objRS93("ASMTDESC") %></td>
    <% End If %>


  </tr>
  <tr valign="top">
    <td class="STitle">HSTD Code/Desc </td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td colspan="2" class="rText2">&nbsp;<%= objRS95("HSTD") %>&nbsp;<%= objRS95("HSTDDESC") %>&nbsp;</td>
    <% End If %>
    	<td align="right" class="STitle"></td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td colspan="2" class="rText2">&nbsp;<%= objRS94("HSTD") %>&nbsp;<%= objRS94("HSTDDESC") %>&nbsp;</td>
    <% End If %>
    	<td class="STitle"></td>
    <% If  objRS93.EOF Then %>
    	<td colspan="2" class="rText2">&nbsp;&nbsp;</td>
    <% Else %>
    	<td colspan="2" class="rText2">&nbsp;<%= objRS93("HSTD") %>&nbsp;<%= objRS93("HSTDDESC") %>&nbsp;</td>
    <% End If %>


  </tr>
  <tr valign="top">
    <td class="STitle">Choice/REL</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td colspan="2" class="rText2">&nbsp;<%= objRS95("HSTDCHC") %>&nbsp;<%= objRS95("HSTDDESC") %>&nbsp;<%= objRS95("RELHSTD") %></td>
    <% End If %>
    <td align="right" class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td colspan="2" class="rText2">&nbsp;<%= objRS94("HSTDCHC") %>&nbsp;<%= objRS94("HSTDDESC") %>&nbsp;<%= objRS95("RELHSTD") %></td>
    <% End If %>
    <td align="right" class="STitle">&nbsp;</td>
    <% If  objRS93.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td colspan="2" class="rText2">&nbsp;<%= objRS93("HSTDCHC") %>&nbsp;<%= objRS93("HSTDDESC") %>&nbsp;<%= objRS95("RELHSTD") %></td>
    <% End If %>

    <td>&nbsp;</td>
  </tr>
  <tr valign="top">
    <td class="STitle">MP #</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td colspan="2" align="center" class="rText2">&nbsp;<%= objRS95("MULTI") %>&nbsp;</td>
    <% End If %>
    <td align="right" class="STitle"></td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td colspan="2" align="center" class="rText2">&nbsp;<%= objRS94("MULTI") %>&nbsp;</td>
    <% End If %>
    <td class="STitle" align="right"></td>
    <% If  objRS93.EOF Then %>
	<td colspan="2" class="rText2">&nbsp;&nbsp;</td>
    <% Else %>
    <td colspan="2" align="center" class="rText2">&nbsp;<%= objRS93("MULTI") %>&nbsp;</td>
    <% End If %>

    <td>&nbsp;</td>
  </tr>
  <tr valign="top">
    <td class="STitle">Land</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td  align="right" class="rText">&nbsp;<font class="oText"><%= objRS95("ESTLAND") %></font></td>
    	<td width="35" align="right" class="rText"></td>
    <% End If %>
    <td class="STitle" align="right"></td>
    <% If  objRS94.EOF Then %>
			<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    <td  align="right" class="rText">&nbsp;<font class="oText"><%= objRS94("ESTLAND") %></font></td>
    <td width="35" align="right" class="rText">&nbsp;</td>
    <% End If %>
    <td align="right" class="STitle">&nbsp;</td>
    <% If  objRS93.EOF Then %>
	<td colspan="2" class="rText">&nbsp;&nbsp;</td>
    <% Else %>
    <td align="right" class="rText">&nbsp;<font class="oText"><%= objRS93("ESTLAND") %></font></td>
    <td class="rText" >&nbsp;&nbsp;</td>
    <% End If %>

    <td>&nbsp;</td>
  </tr>
  <tr valign="top">
    <td class="STitle">G A Land</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td align="right"  class="rText2"><%= objRS95("ESTGA") %></td>
        <td align="right"  class="rText2">&nbsp;</td>
    <% End If %>
    <td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    <td class="rText2" align="right" ><%= objRS94("ESTGA") %></td>
    <td align="right" valign="top" class="rText2"></td>
    <% End If %>
    <td align="right" class="STitle">&nbsp;</td>
    <% If  objRS93.EOF Then %>
	<td colspan="2" class="rText2">&nbsp;&nbsp;</td>
    <% Else %>
    <td class="rText2" align="right" ><%= objRS93("ESTGA") %></td>
    <td class="rText2" >&nbsp;&nbsp;</td>
    <% End If %>

    <td>&nbsp;</td>
  </tr>
  <tr valign="top">
    <td class="STitle">Building</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td align="right" class="rText"><%= objRS95("ESTbuild") %></td>
    	<td align="right"  class="rText">&nbsp;</td>
    <% End If %>
    <td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td align="right" class="rText"><%= objRS94("ESTbuild") %></td>
    	<td class="rText" align="right">&nbsp;</td>
    <% End If %>
    	<td class="STitle" align="right"></td>
    <% If  objRS93.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
    <% Else %>
    	<td align="right" class="rText"><%= objRS93("ESTbuild") %></td>
    	<td class="rText" >&nbsp;&nbsp;</td>
    <% End If %>

    <td>&nbsp;</td>
  </tr>
  <tr valign="top">
    <td  class="STitle">Machine</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td  align="right" class="rText2">&nbsp;<font class="oText"><%= objRS95("ESTmach") %></font></td>
		<td align="right" class="rText2"></td>
	<% End If %>
    <td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    <td align="right" class="rText2">&nbsp;<font class="oText"><%= objRS94("ESTmach") %></font></td>
    <td class="rText2" align="right"></td>
    <% End If %>
    <td class="STitle">&nbsp;</td>
    <% If  objRS93.EOF Then %>
	<td colspan="2" class="rText2">&nbsp;&nbsp;</td>
    <% Else %>
    <td align="right" class="rText2">&nbsp;<font class="oText"><%= objRS93("ESTmach") %></font></td>
    <td class="rText2" >&nbsp;&nbsp;</td>
    <% End If %>

    <td>&nbsp;</td>
  </tr>
  <tr valign="top">
    <td  class="STitle">Total Market</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    <td align="right" class="rText">&nbsp;<%= objRS95("ESTtotal") %></td>
    <td align="right" class="rText"></td>
    <% End If %>
    <td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td align="right" class="rText">&nbsp;<%= objRS94("ESTtotal") %></td>
    	<td align="right" class="rText"></td>
    <% End If %>
    <td class="STitle">&nbsp;</td>
    <% If  objRS93.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
    <% Else %>
    	<td align="right" class="rText">&nbsp;<%= objRS93("ESTtotal") %></td>
    	<td class="rText" >&nbsp;&nbsp;</td>
    <% End If %>

  </tr>
  <tr valign="top">
    <td  class="STitle">Total Tax Mrkt</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td align="right" class="rText2"><%= objRS95("TAXtotal") %></td>
		<td align="right"  class="rText2">&nbsp;</td>
	<% End If %>
	<td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
		<td align="right" class="rText2"><%= objRS94("TAXtotal") %></td>
		<td class="rText2" align="right">&nbsp;</td>
	<% End If %>
	<td class="STitle" align="right"></td>
	<% If  objRS93.EOF Then %>
		<td colspan="2" class="rText2">&nbsp;&nbsp;</td>
    <% Else %>
		<td align="right" class="rText2"><%= objRS93("TAXtotal") %></td>
		<td class="rText2" >&nbsp;&nbsp;</td>
	<% End If %>
    <td>&nbsp;</td>
  </tr>
  <tr valign="top">
    <td  class="STitle">Net T C</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    <td align="right" class="rText">&nbsp;<%= objRS95("Netcap") %></td>
    <td align="right" class="rText"></td>
    <% End If %>
    <td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    <td align="right" class="rText">&nbsp;<%= objRS94("Netcap") %></td>
    <td align="right" class="rText"></td>
    <% End If %>
    <td class="STitle">&nbsp;</td>
    <% If  objRS93.EOF Then %>
	<td colspan="2" class="rText">&nbsp;&nbsp;</td>
    <% Else %>
    <td align="right" class="rText">&nbsp;<%= objRS93("Netcap") %></td>
    <td class="rText">&nbsp;</td>
    <% End If %>

  </tr>
  <tr valign="top">
    <td class="STitle">Bldg Site</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td align="right" class="rText2">&nbsp;<font class="oText"><%= objRS95("Estsite") %></font></td>
		<td align="right"  class="rText2">&nbsp;</td>
	<% End If %>
	<td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
		<td align="right" class="rText2"><font class="oText"><%= objRS94("Estsite") %></font></td>
		<td class="rText2" align="right">&nbsp;</td>
	<% End If %>
	<td class="STitle" align="right"></td>
	<% If  objRS93.EOF Then %>
	<td colspan="2" class="rText2">&nbsp;&nbsp;</td>
    <% Else %>
	<td align="right" class="rText2"><font class="oText"><%= objRS93("Estsite") %></font></td>
	<td class="rText2" >&nbsp;&nbsp;</td>
	<% End If %>


  </tr>
  <tr valign="top">
    <td  class="STitle">Till Land</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td align="right" class="rText">&nbsp;<%= objRS95("Till") %></td>
		<td align="right" class="rText"></td>
	<% End If %>
	<td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
		<td align="right" class="rText">&nbsp;<%= objRS94("Till") %></td>
		<td align="right" class="rText"></td>
	<% End If %>
	<td class="STitle">&nbsp;</td>
	<% If  objRS93.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
    <% Else %>
		<td align="right" class="rText">&nbsp;<%= objRS93("Till") %></td>
		<td class="rText">&nbsp;</td>
	<% End If %>

  </tr>
  <tr valign="top">
    <td  class="STitle">New Improve</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td align="right" class="rText2">&nbsp;<font class="oText"><%= objRS95("Improve") %></font></td>
		<td align="right"  class="rText2">&nbsp;</td>
	<% End If %>
	<td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
		<td align="right" class="rText2">&nbsp;&nbsp;<font class="oText"><%= objRS94("Improve") %></font></td>
		<td class="rText2" align="right">&nbsp;</td>
	<% End If %>
	<td class="STitle" align="right"></td>
	<% If  objRS93.EOF Then %>
		<td colspan="2" class="rText2">&nbsp;&nbsp;</td>
    <% Else %>
		<td align="right" class="rText2">&nbsp;&nbsp;<font class="oText"><%= objRS93("Improve") %></font></td>
		<td class="rText2" >&nbsp;&nbsp;</td>
	<% End If %>
  </tr>
  <tr valign="top">
    <td  class="STitle">Deeded Acres</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    <td align="right" class="rText"><%= objRS95("Deed") %></td>
    <td align="right" class="rText"></td>
    <% End If %>
	<td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    <td align="right" class="rText"><%= objRS94("Deed") %></td>
    <td align="right" class="rText"></td>
    <% End If %>
	<td class="STitle">&nbsp;</td>
	<% If  objRS93.EOF Then %>
	<td colspan="2" class="rText">&nbsp;&nbsp;</td>
    <% Else %>
    <td align="right" class="rText"><%= objRS93("Deed") %></td>
    <td class="rText">&nbsp;</td>
    <% End If %>

  </tr>
  <tr valign="top">
    <td  class="STitle">Tillable Acres</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
	    <td align="right" class="rText2"><%= objRS95("Tillacres") %></td>
		<td align="right"  class="rText2">&nbsp;</td>
    <% End If %>
	<td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
		<td align="right" class="rText2"><%= objRS94("Tillacres") %></td>
		<td class="rText2" align="right">&nbsp;</td>
	<% End If %>
	<td class="STitle" align="right"></td>
	<% If  objRS93.EOF Then %>
		<td colspan="2" class="rText2">&nbsp;&nbsp;</td>
    <% Else %>
		<td align="right" class="rText2"><%= objRS93("Tillacres") %></td>
		<td class="rText2" >&nbsp;&nbsp;</td>
	<% End If %>
  </tr>
  <tr valign="top">
    <td  class="STitle">House/Garage</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td align="right" valign="top" class="rText">&nbsp;<font class="oText"><%= objRS95("Esthouse") %></font></td>
		<td align="right"  class="rText">&nbsp;</td>
	<% End If %>
	<td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
		<td align="right" class="rText">&nbsp;&nbsp;<font class="oText"><%= objRS94("Esthouse") %></font></td>
		<td class="rText" align="right">&nbsp;</td>
	<% End If %>
	<td class="STitle" align="right"></td>
	<% If  objRS93.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
    <% Else %>
		<td align="right" class="rText">&nbsp;&nbsp;<font class="oText"><%= objRS93("Esthouse") %></font></td>
		<td class="rText" >&nbsp;&nbsp;</td>
	<% End If %>

  </tr>
  <tr valign="top">
    <td class="STitle">Other Bldg</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td align="right" class="rText2"><%= objRS95("Estother") %></td>
		<td align="right" class="rText2"></td>
	<% End If %>
	<td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
		<td align="right" class="rText2"><%= objRS94("Estother") %></td>
		<td align="right" class="rText2"></td>
	<% End If %>
	<td class="STitle">&nbsp;</td>
	<% If  objRS93.EOF Then %>
		<td colspan="2" class="rText2">&nbsp;&nbsp;</td>
    <% Else %>
		<td align="right" class="rText2"><%= objRS93("Estother") %></td>
		<td class="rText2">&nbsp;</td>
	<% End If %>

  </tr>
  <tr valign="top">
    <td class="STitle">Limit Flag</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td align="right" valign="top" class="rText">&nbsp;<font class="oText"><%= objRS95("Limited") %></font></td>
		<td align="right"  class="rText">&nbsp;</td>
	<% End If %>
	<td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
		<td align="right" class="rText">&nbsp;&nbsp;<font class="oText"><%= objRS94("Limited") %></font></td>
		<td class="rText" align="right">&nbsp;</td>
	<% End If %>
	<td class="STitle" align="right"></td>
	<% If  objRS93.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
    <% Else %>
		<td align="right" class="rText">&nbsp;&nbsp;<font class="oText"><%= objRS93("Limited") %></font></td>
		<td class="rText" >&nbsp;&nbsp;</td>
	<% End If %>


  </tr>
  <tr valign="top">
    <td class="STitle">Yr Appraised</td>
    <% If  objRS95.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
    	<td align="right" class="rText2"><%= objRS95("Yrappr") %></td>
		<td align="right" class="rText2"></td>
	<% End If %>
	<td class="STitle">&nbsp;</td>
    <% If  objRS94.EOF Then %>
		<td colspan="2" class="rText">&nbsp;&nbsp;</td>
	<% Else %>
		<td align="right" class="rText2"><%= objRS94("Yrappr") %></td>
		<td align="right" class="rText2"></td>
	<% End If %>
	<td class="STitle">&nbsp;</td>
	<% If  objRS93.EOF Then %>
		<td colspan="2" class="rText2">&nbsp;&nbsp;</td>
    <% Else %>
		<td align="right" class="rText2"><%= objRS93("Yrappr") %></td>
		<td class="rText2">&nbsp;</td>
	<% End If %>

  </tr>
</table>

<br>
<%
Dim counter, intTotalRec
intTotalRec = intI5
'Response.Write("intTotalRec= ") & intTotalRec
If intTotalRec <> 1   Then

'Response.Write("<A HREF='Parcel.asp?pid=" & sesParcel & "&tid=5'><IMG SRC='Images/nextrecord.gif' ALT='Next Record' BORDER=0></A>")

   End if
'counter = request.form("first")
'counter = counter + 1
'Response.write("counter = ") & counter
'Response.write("intI5= ") & intI5
'Session("counter")  = counter
sesRecNum = Session("counter")
'response.Write("sesrecnum#2 = ") & sesRecNum
'response.write("session(counter) = ") & Session("counter")
'Response.Write("intTotalRec= ") & intTotalRec
'Response.Write("sesTotalRec= ") & sesTotalRec



%>
<%
'Response.Write("sesRecNum= ") & sesRecNum
'Response.Write("intI5= ") & intI5
If intI5 = 1 Then
Else
If sesRecNum <= intI5 Then

%>
<form action="" method="post">
<center>
<input type="hidden" name="first" value="<%=counter%>">


<input name="decisionButton" type="submit" value="Next Record">&nbsp;&nbsp;&nbsp;
</center>
</form>
<% Else
counter = 0
%>

<%
End If
End If


	If RecNumCase => intTotalRec Then
		count = 1
		Session("counter") = count
		sesRecNum = Session("counter")
		sesRecNum = 1
		RecNumCase = sesRecNum
'		Response.Write("sesRecNum#11= ") & sesRecNum
'		Response.Write("RecNumCase#11= ") & RecNumCase
	End If


%>
<br>
<br><br>
