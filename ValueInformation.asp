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
<%' check to see how many records are in table 9 for each parcel #
	Session("TOTREC") = objRSV2("RecNum")
%>
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
	<tr>
		<td ></td>
		<td ></td>
		<td ></td>
		<td <STRONG><font color="#FF0000"><%= rs("DELETED") %></font></STRONG></td>

	</tr>
</table>

<br><br>
<table border="0" cellpadding="0" cellspacing="0">
  <tr valign="top">
    <td width="120" class="SHeader">True/ Full Values</td>
    <td width="80" class="SHeader"></td>
    <td width="20" class="SHeader"></td>
    <td width="110" class="SHeader">Assessed Values</td>
    <td width="80" class="SHeader"></td>
    <td width="20" class="SHeader"></td>
    <td width="110" class="SHeader">Miscellaneous</td>
    <td width="60" class="SHeader"></td>
  </tr>
  <tr valign="top">
    <td colspan="8"></td>
  </tr>
  <tr valign="top">
    <td class="STitle">Land</td>
    <td align="right" class="rText"><%= FormatNumber(rs("LAND"), 0) %></td>
    <td class="STitle"></td>
    <td class="STitle">Assessed</td>
    <td align="right" class="rText"><%= FormatNumber(rs("ASSESSED"), 0) %></td>
    <td class="STitle"></td>
    <td class="STitle">ASMT</td>
    <td align="right" class="rText"><%= FormatNumber(rs("ASMT"), 0) %></td>
  </tr>
  <tr valign="top">
    <td class="STitle">Building</td>
    <td align="right" class="rText2"><%= FormatNumber(rs("Bldg"), 0) %></td>
    <td class="STitle"></td>
    <td class="STitle">Taxable</td>
    <td align="right" class="rText2"><%= FormatNumber(rs("Taxable"), 0) %></td>
    <td class="STitle"></td>
    <td class="STitle">HSTD</td>
    <td  align="right" class="rText2"><%= rs("ASMTDESC") %></td>
  </tr>
  <tr valign="top">
    <td class="STitle">Total</td>
    <td align="right" class="rText"><%= FormatNumber(rs("TOTVAL"), 0) %></td>
    <td class="STitle"></td>
    <td class="STitle">HSTD Credit</td>
    <td align="right" class="rText"><%= FormatNumber(rs("HstdCd"), 0) %></td>
    <td class="STitle"></td>
    <td class="STitle">Deeded Acres</td>
    <td align="right" class="rText"><%= rs("DeedAc") %></td>
  </tr>
  <tr valign="top">
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle">Net Taxable</td>
    <td align="right" class="rText2"><%= FormatNumber(rs("NetTax"), 0) %></td>
    <td class="STitle"></td>
    <td class='STitle'>Tillable Acres</td>
    <td align='right' class='rText2'><%= rs("TillAC") %></td>


  </tr>
  <tr valign="top">
    <td class="STitle">Till Land</td>
    <td align="right" class="rText"><%= rs("TILL") %></td>
    <td class="STitle"></td>
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
  </tr>
  <tr valign="top">
    <td class="STitle">&nbsp</td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
  </tr>
  <tr valign="top">
    <td class="STitle">&nbsp</td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
  </tr>
  <tr valign="top">
    <td class="STitle">&nbsp</td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
  </tr>
  <tr valign="top">
    <td class="STitle">&nbsp</td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
  </tr>
  <tr valign="top">
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="SHeader">Lot Dimension</td>
    <td align="right" class="STitle"></td>
  </tr>
  <tr valign="top">
    <td class="STitle"></td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
   	<td class="STitle"></td>
   	<td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle">Zoning Code</td>
    <td align="right" class="rText2"><%= rs("ZONE") %></td>
  </tr>
  <tr valign="top">
    <td class="sTitle"></td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle"></td>
	<td align="right" class="STitle"></td>
	<td class="STitle"></td>
    <td class="STitle">Lot Dimension</td>
	<% If rs("LOT1") <> 0 Then %>
	    <td align="right" class="rText"><%= rs("LOT1") %>&nbsp;X&nbsp;<%= rs("LOT2") %></td>
	<% Else %>
	    <td align="right" class="rText"></td>
	<% End If %>

  </tr>
    <tr valign="top">
      <td class="STitle"></td>
      <td align="right" class="STitle"></td>
      <td class="STitle"></td>
     	<td class="STitle"></td>
     	<td align="right" class="STitle"></td>
      <td class="STitle"></td>
      <td class="STitle">Square Footage</td>
      <td align="right" class="rText2"><%= rs("FOOT") %></td>
  </tr>
  <tr valign="top">
    <td class="sTitle"></td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="sTitle"></td>
    <td align="right" class="STitle"></td>
    <td class="STitle"></td>
    <td class="STitle">Calc Units</td>
    <td align="right" class="rText"><%= rs("UNITS") %></td>
  </tr>

  </table>

<table border="0" cellpadding="0" cellspacing="0" width="600" class="sTitle">
<tr valign="top">
    <td class="tText2" align="center">This Data is Subject to Change</td>
  </tr>
</table>