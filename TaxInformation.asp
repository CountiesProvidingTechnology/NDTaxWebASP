

<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="80" class="SHeader">General</td>
		<td width="200" class="SHeader"></td>
		<td width="70" class="SHeader"></td>
		<td width="250" class="SHeader"></td>
	</tr>
	<tr valign="top">
		<td colspan="4" height="1" bgcolor="#000000"></td>
	</tr>
	<tr valign="top">
		<td class="STitle">Receipt #</td>
		<td class="rText"><%= objRS("TXRCPT#") %></td>
		<td class="STitle">Name</td>
		<td class="rText"><%= objRS("TPNREV") %></td>
	</tr>
	<tr>
		<td class="STitle">ASMT</td>
		<td class="rText"><%= objRS("TXASM1") %>&nbsp;<%= objRS("TXASMD") %></td>
		<td class="STitle">MP#</td>
		<td class="rText"><%= objRS("TXMP#") %></td>
	</tr>
	<tr>
		<td class="STitle">Homestead</td>
		 <% If Trim(objRS("TXHSTP")) <> ""  Then %>
		 	<td class="rText"><%= objRS("TXHSTC") %>&nbsp;<%= objRS("TXHSTD") %></td>

		 <% Else %>
			<td class="rText"></td>
    	<% End If %>
		<td class="STitle">MP Name</td>
		<td class="rText"><%= objRS("MPNAME") %></td>
	</tr>
	<tr>
		<td class="STitle">HS Percent</td>
		<% If Trim(objRS("TXHSTP")) <> ""  Then %>
			<td class="rText"><%= objRS("TXHSTP") %></td>

		<% Else %>
			<td class="rText"></td>
		<% End If %>
		<td class="STitle"></td>
		<td class="rText"></td>
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
    <td width="120" class="SHeader">Market/Tax</td>
    <td width="60" class="SHeader"></td>
    <td width="40" class="SHeader"></td>
    <td width="110" class="SHeader"></td>
    <td width="60" class="SHeader"></td>
    <td width="40" class="SHeader"></td>
    <td width="110" class="SHeader"></td>
    <td width="60" class="SHeader"></td>
  </tr>
  <tr valign="top">
    <td colspan="8"></td>
  </tr>
  <tr valign="top">
    <td class="STitle">T & F Land </td>
    <td align="right" class="rText"><%= FormatNumber(objRS("TPLAND"), 0) %></td>
    <td class="rText"></td>
    <td class="STitle">Tax State</td>
    <td align="right" class="rText"><%= FormatNumber(objRS("TPTXST"), 2) %></td>
    <td class="rText"></td>
    <td class="STitle">Gross Tax</td>
    <td align="right" class="rText"><%= FormatNumber(objRS("GROSSTAX"), 2) %></td>
  </tr>
  <tr valign="top">
    <td class="STitle">T & F Building</td>
    <td align="right" class="rText2"><%= FormatNumber(objRS("TPBLDG"), 0) %></td>
    <td class="rText2"></td>
    <td class="STitle">Tax County</td>
    <td align="right" class="rText2"><%= FormatNumber(objRS("TPCNTY"), 2) %></td>
    <td class="rText2"></td>
    <td class="STitle">ST PD Cred</td>
    <td  align="right" class="rText2"><%= FormatNumber(objRS("OPEN1"), 2) %></td>
  </tr>
  <tr valign="top">
    <td class="STitle">Total T & F</td>
    <td align="right" class="rText"><%= FormatNumber(objRS("TPLAND") + objRS("TPBLDG"), 0) %></td>
    <td class="rText"></td>
    <td class="STitle">Tax Twp/Cty</td>
    <td align="right" class="rText"><%= FormatNumber(objRS("TPTOWN"), 2) %></td>
    <td class="rText"></td>
    <td class="STitle">Special Asmt</td>
    <td  align="right" class="rText"><%= FormatNumber(objRS("TPTASM"), 2) %></td>
  </tr>
  <tr valign="top">
    <td class="STitle">Assessed</td>
    <td align="right" class="rText2"><%= FormatNumber(objRS("TPAVGR"), 0) %></td>
    <td class="rText2"></td>
    <td class="STitle">Tax School</td>
    <td align="right" class="rText2"><%= FormatNumber(objRS("TPSKOL"), 2) %></td>
    <td class="rText2"></td>
    <td class="STitle">Tax Due</td>
    <td align="right" class="rText2"><%= FormatNumber(objRS("TAXDUE"), 2) %></td>
  </tr>
  <tr valign="top">
    <td class="STitle">Taxable</td>
    <td align="right" class="rText"><%= FormatNumber(objRS("TPAVNT"), 0) %></td>
    <td class="rText"></td>
    <td class="STitle">Tax Cnty.WD</td>
    <td align="right" class="rText"><%= FormatNumber(objRS("TPWIDE"), 2) %></td>
    <td class="rText"></td>
    <td class="STitle"></td>
    <td align="right" class="rText"></td>
  </tr>
  <tr valign="top">
    <td class="STitle">HSTD Credit</td>
    <td align="right" class="rText2"><%= FormatNumber(objRS("TPAVHC"), 0) %></td>
    <td class="rText2"></td>
    <td class="STitle">Tax Increm</td>
    <td align="right" class="rText2"><%= FormatNumber(objRS("TPTINT"), 2) %></td>
    <td class="rText2"></td>
    <td class="STitle">Disc Avail</td>
    <td align="right" td class="rText2"><%= FormatNumber(objRS("TPDISA"), 2) %></td>
    <td align="right" class="rText"></td>
  </tr>
  <tr valign="top">
  <td class="STitle">Net Taxable</td>
    <td align="right" class="rText"><%= FormatNumber(objRS("TPAVNT"), 0) %></td>
    <td class="rText"></td>
        <% If cid = 02 Then %>
			<td class="STitle">Fire</td>
			<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC1"), 2) %></td>
    	<% End If %>
        <% If cid = 23 Then %>
			<td class="STitle">Fire</td>
			<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC1"), 2) %></td>
    	<% End If %>
        <% If cid = 27 Then %>
			<td class="STitle">Fire</td>
			<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC1"), 2) %></td>
    	<% End If %>
        <% If cid = 47 Then %>
			<td class="STitle">Water</td>
			<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC1"), 2) %></td>
    	<% End If %>
        <% If cid = 13 Then %>
			<td class="STitle">Fire</td>
			<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC1"), 2) %></td>
    	<% End If %>
        <% If cid = 30 Then %>
			<td class="STitle">Fire</td>
			<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC1"), 2) %></td>
    	<% End If %>
        <% If cid = 31 Then %>
			<td class="STitle">Fire</td>
			<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC1"), 2) %></td>
    	<% End If %>
        <% If cid = 37 Then %>
			<td class="STitle">Fire</td>
			<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC1"), 2) %></td>
    	<% End If %>
        <% If cid = 34 Then %>
			<td class="STitle">Fire</td>
			<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC1"), 2) %></td>
    	<% End If %>
        <% If cid = 41 Then %>
			<td class="STitle">Fire</td>
			<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC1"), 2) %></td>
    	<% End If %>

    <td class="rText"></td>
    <td class="STitle">Net Tax Due</td>
    <td align="right" class="rText"><%= FormatNumber(objRS("NETDUE"), 2) %></td>
  </tr>
  <tr valign="top">
    <td class="STitle">Mill Rate</td>
    <td align="right" class="rText2"><%= (objRS("TPTMR")) * 1000 %></td>
    <td class="rText2"></td>
        <% If cid = 02 Then %>
		   	<td class="STitle">Park</td>
		   	<td align="right" class="rText2"><%= Formatnumber(objRS("TPSPC2"), 2) %></td>
    	<% End If %>
        <% If cid = 23 Then %>
		   	<td class="STitle">Park</td>
		   	<td align="right" class="rText2"><%= Formatnumber(objRS("TPSPC2"), 2) %></td>
    	<% End If %>
        <% If cid = 27 Then %>
		   	<td class="STitle">Soil</td>
		   	<td align="right" class="rText2"><%= Formatnumber(objRS("TPSPC2"), 2) %></td>
    	<% End If %>
        <% If cid = 47 Then %>
		   	<td class="STitle"></td>
		   	<td align="right" class="rText2"></td>
    	<% End If %>
        <% If cid = 13 Then %>
		   	<td class="STitle">Ambl</td>
		   	<td align="right" class="rText2"><%= Formatnumber(objRS("TPSPC2"), 2) %></td>
    	<% End If %>
        <% If cid = 30 Then %>
		   	<td class="STitle">Park</td>
		   	<td align="right" class="rText2"><%= Formatnumber(objRS("TPSPC2"), 2) %></td>
    	<% End If %>
        <% If cid = 37 Then %>
		   	<td class="STitle">Park</td>
		   	<td align="right" class="rText2"><%= Formatnumber(objRS("TPSPC2"), 2) %></td>
    	<% End If %>
        <% If cid = 31 Then %>
		   	<td class="STitle">Misc</td>
		   	<td align="right" class="rText2"><%= Formatnumber(objRS("TPSPC2"), 2) %></td>
    	<% End If %>
        <% If cid = 34 Then %>
		   	<td class="STitle">Debt</td>
		   	<td align="right" class="rText2"><%= Formatnumber(objRS("TPSPC2"), 2) %></td>
    	<% End If %>
        <% If cid = 41 Then %>
		   	<td class="STitle">Park</td>
		   	<td align="right" class="rText2"><%= Formatnumber(objRS("TPSPC2"), 2) %></td>
    	<% End If %>
    <td class="rText2"></td>
    <td class="STitle"></td>
    <td class="rText2"></td>
  </tr>
  <tr valign="top">
    <td class="STitle">Statement #</td>
    <td align="right" class="rText"><%= objRS("TXRCPT#") %></td>
    <td class="rText"></td>
    <% If cid = 27 Then %>
	   	<td class="STitle">Misc</td>
	   	<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC3"), 2) %></td>
    <% End If %>
    <% If cid = 47 Then %>
	   	<td class="STitle">Fire</td>
	   	<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC3"), 2) %></td>
    <% End If %>
    <% If cid = 02 Then %>
	   	<td class="STitle"></td>
	   	<td align="right" class="rText"></td>
    <% End If %>
    <% If cid = 23 Then %>
	   	<td class="STitle"></td>
	   	<td align="right" class="rText"></td>
    <% End If %>
    <% If cid = 13 Then %>
	   	<td class="STitle"></td>
	   	<td align="right" class="rText"></td>
    <% End If %>
    <% If cid = 30 Then %>
	   	<td class="STitle">Water</td>
	   	<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC3"), 2) %></td>
    <% End If %>
    <% If cid = 31 Then %>
	   	<td class="STitle">AMB</td>
	   	<td align="right" class="rText"><%= Formatnumber(objRS("TPSPC3"), 2) %></td>
    <% End If %>
    <% If cid = 34 Then %>
	   	<td class="STitle"></td>
	   	<td align="right" class="rText"></td>
    <% End If %>
    <% If cid = 37 Then %>
	   	<td class="STitle"></td>
	   	<td align="right" class="rText"></td>
    <% End If %>
    <% If cid = 41 Then %>
	   	<td class="STitle"></td>
	   	<td align="right" class="rText"></td>
    <% End If %>
    <td class="rText"></td>
    <td class="STitle">Tax AB/Adds</td>
    <td align="right" class="rText"><%= FormatNumber(objRS("TPTAB"), 2) %></td>
  </tr>
  <tr valign="top">
    <td class="STitle"></td>
    <td align="right" class="rText2"></td>
    <td class="rText2"></td>
    	<% If cid = 27 Then %>
			   	<td class="STitle">COMR</td>
			   	<td align="right" class="rText2"><%= Formatnumber(objRS("TPSPC4"), 2) %></td>
    	<% End If %>
    	<% If cid = 02 Then %>
			   	<td class="STitle"></td>
			   	<td align="right" class="rText2"></td>
    	<% End If %>
    	<% If cid = 23 Then %>
			   	<td class="STitle"></td>
			   	<td align="right" class="rText2"></td>
    	<% End If %>
    	<% If cid = 47 Then %>
			   	<td class="STitle"></td>
			   	<td align="right" class="rText2"></td>
    	<% End If %>
    	<% If cid = 13 Then %>
			   	<td class="STitle"></td>
			   	<td align="right" class="rText2"></td>
    	<% End If %>
    	<% If cid = 30 Then %>
			   	<td class="STitle"></td>
			   	<td align="right" class="rText2"></td>
    	<% End If %>
    	<% If cid = 31 Then %>
			   	<td class="STitle">Soil</td>
			   	<td align="right" class="rText2"><%= Formatnumber(objRS("TPSPC4"), 2) %></td>
    	<% End If %>
    	<% If cid = 34 Then %>
			   	<td class="STitle"></td>
			   	<td align="right" class="rText2"></td>
    	<% End If %>
    	<% If cid = 37 Then %>
			   	<td class="STitle"></td>
			   	<td align="right" class="rText2"></td>
    	<% End If %>
    	<% If cid = 41 Then %>
			   	<td class="STitle"></td>
			   	<td align="right" class="rText2"></td>
    	<% End If %>
    <td class="rText2"></td>
    <td class="STitle">S A AB/Adds</td>
    <td align="right" class="rText2"><%= FormatNumber(objRS("TPSAAB"), 2) %></td>
  </tr>

  <tr valign="top">
    <td class="STitle"></td>
    <td align="right" class="rText"></td>
    <td class="rText"></td>

		   	<td class="STitle">Tax Penalty</td>
		   	<td align="right" class="rText"><%= FormatNumber(objRS("TPPDUE"), 2) %></td>

    <td class="rText"></td>
    <td class="STitle"></td>
    <td align="right" class="rText"></td>
  </tr>
  <tr valign="top">
    <td class="sTitle"></td>
    <td align="right" class="rText2"></td>
    <td class="rText2"></td>

		   	<td class="STitle">Tax Interest</td>
		   	<td align="right" class="rText2"><%= FormatNumber(objRS("TPINTR"), 2) %></td>

    <td class="rText2"></td>
    <td class="STitle">Adj.NT.Due</td>
    <td  align="right" class="rText2"><%= FormatNumber(objRS("ADJNET"), 2) %></td>
  </tr>
  <tr valign="top">
    <td class="STitle"></td>
    <td align="right" class="rText"></td>
    <td class="rText"></td>
    <td class="STitle">SA Penalty</td>
    <td align="right" class="rText"><%= FormatNumber(objRS("TPSAPN"), 2) %></td>
    <td  class="rText"></td>
    <td class="sTitle">Total Receipts</td>
    <td align="right" class="rText"><%= FormatNumber(objRS("TPTRC"), 2) %></td>
    <td class="rText"></td>
    <td align="right" class="rText"></td>
  </tr>
  <tr valign="top">
    <td class="sTitle"></td>
    <td align="right" class="rText2"></td>
    <td class="rText2"></td>
    <td class="STitle">SA Interest</td>
    <td align="right" class="rText2"><%= FormatNumber(objRS("TPSAIN"), 2) %></td>
    <td class="rText2"></td>
    <td class="STitle">Disc Taken</td>
    <td align="right" class="rText2"><%= FormatNumber(objRS("TPDISC"), 2) %></td>
  </tr>
  <tr valign="top">
    <td class="rText"></td>
    <td align="right" class="rText"></td>
    <td class="rText"></td>
    <td class="STitle">Cost</td>
    <td align="right" class="rText"></td>
    <td class="rText"></td>
    <td class="STitle">Remain Due</td>
    <td align="right" class="rText"><%= FormatNumber(objRS("REMDUE"), 2) %></td>
  </tr>
  <tr>
    <td class="rText2"></td>
    <td class="rText2"></td>
    <td class="rText2"></td>
    <td class="STitle"></td>
    <td class="rText2" align="right"></td>
    <td class="rText2"></td>
    <td class="rText2"></td>
    <td class="rText2"></td>
  </tr>
  <tr>
    <td class="STitle" ></td>
    <td class="rText" align="right"></td>
    <td class="rText"></td>
    <td class="STitle"></td>
    <td class="rText" align="right"></td>
    <td class="rText"></td>
    <td class="STitle"></td>
    <td class="rText"  align="right"></td>
  </tr>

</table>