<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Notice of Disclaimer</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
	Dim cid, cName
	cid = request.QueryString("cid")

	If cid = "" Then
		response.Redirect("tax.asp")
	End If

	Select Case cid
		Case 30
		cName = "Morton"
		Case 47
		cName = "Stutsman"
	End Select

	Session("CountyID")= cid
	Session("County Name")= cName

	response.Write("<link rel='stylesheet' href='" & cid & ".css' type='text/css'>")
%>
</head>

<body>

<!-- #include file="PageCounters.asp" -->

<p>


	<table border="0" cellspacing="0" cellpadding="0" align="center">
		<tr valign="top">
			<td width="450" align="justify" class="STitle">Notice of Disclaimer:</td>
		</tr>
		<tr>
			<td width="450" bgcolor="#000000"></td>
		</tr>
		<tr>
			<td width="450" align="justify" class="rText">
			<% Response.Write(cName & " County makes the Web information available on an &quot;as is&quot; basis.  All warranties and representations of any kind with regard to said information is disclaimed, including the implied warranties of merchantability and fitness for a particular use.  " & cName & " County does not warrant the information against deficiencies of any kind.  Under no circumstances will " & cName & " County, or any of its officers or employees be liable for any consequential, incidental, special or exemplary damages even if appraised of the likelihood of such damages occurring.</td>") %>
		</tr>
	</table>
<form action="Searchinput.asp?" method="post">
<center>
<input type="radio" name="decision" value="iAccept">I ACCEPT the terms and conditions provided within this disclaimer.<br>
<input type="radio" name="decision" value="iDecline" checked>I DO NOT AGREE to the terms and conditions provided here.<br><br>
<input type="hidden" name="cid" value=<%=cid%> >
<input name="decisionButton" type="submit" value="Proceed">
</center>

</form>
<center>

</center>
</body>
</html>