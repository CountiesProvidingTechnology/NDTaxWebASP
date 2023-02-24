<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>County Parcel Search Page</title>
<%
	Dim cid, cName
	cid = Session("CountyID")
	cName = Session("County Name")
	response.Write("<link rel='stylesheet' href='" & cid & ".css' type='text/css'>")
%>
</head>
<!-- #include file="insDB.asp" -->
<body text="#000000" bgcolor="#FFFFFF">
<%
	If cid = "" Then
		response.Redirect("tax.asp")
	End If

	Dim strResponse, strSearchAgain
	strResponse = Request.Form("decision")
	strSearchAgain = request.QueryString("Again")

	If strResponse = "iDecline" Then
		response.Redirect("Decline.asp")
	Else
		If strResponse = "" AND strSearchAgain = "" Then
			response.Redirect("Disclaimer_Rock.asp")
		End If
	End If

	Dim objCommand, objRS, strQueryString, strPID, strTID, objRS3, objRS5

	Set objCommand = Server.CreateObject("ADODB.Command")

	'strPID = request.QueryString("pid")
	'strTID = request.QueryString("tid")
	'strQueryString = request.QueryString("pid")
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strQueryString & "'"

	'Fill in the command properties'
	objCommand.ActiveConnection = strConnect
	objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] "
	objCommand.CommandType = 1

	Set objRS = objCommand.Execute

	Set objCommand = Nothing


	If cName = "Douglas" Then
		Response.Write("<img src='Images/depart23.gif'>")
	End If
%>

<form action="ParcelList_Rock.asp" method="post">
<table border="0" cellspacing="0" cellpadding="0">
	<tr valign="top">
		<td width="100"></td>
		<td width="550">
			<%
				If cid = 54 or cid = 34 or cid = 67 or cid = 87 or cid = 65 or cid = 53 Then
			%>
			<!-- #include file="SearchByName.asp" -->
			<%
				End If
			%>
				<br>
				<br>
			<%
				If cid = 54 or cid = 34 or cid = 21 or cid = 65 or cid = 61 or cid = 67 or cid = 75 or cid = 87 or cid = 76 or cid = 53 or cid = 47 Then
			%>
			<!-- #include file="SearchByParcel.asp" -->
				<br>
				<br>
			<!-- #include file="SearchByE911.asp" -->
				<br>
				<br>
			<!-- #include file="SearchBySectTwpRange.asp" -->
			<%
				End If
			%>
				<br>
				<br>
			<table border="0" cellspacing="0" cellpadding="0">
				<tr valign="top">
					<td width="450" align="justify" class="STitle">Search Examples and Instructions</td>
				</tr>
				<tr>
					<td width="450" bgcolor="#000000"></td>
				</tr>
				<tr>
					<td width="450" class="rText">The use of a wild card is permitted.  The wild card for this search is '*'.  <br>For example:<br><b>03*</b> in the Parcel Search field would result in a list of all parcels that begin with 03.<br>or<br>
              <b>*25*</b> in the E911 Address Search field would produce a list
              of all parcels that have a 25 within their address.</td>
				</tr>
			</table>
			<br>
			<br>
			<table border="0" cellspacing="0" cellpadding="0">
				<tr valign="top">
					<td width="450" align="justify" class="STitle">Notice of Disclaimer:</td>
				</tr>
				<tr>
					<td width="450" bgcolor="#000000"></td>
				</tr>
				<tr>
					<td width="450" align="justify" class="rText">
					<% Response.Write(cName & " County makes the Web information available on an &quot;as is&quot; basis.  All warranties and representations of any kind with regard to said information is disclaimed, including the implied warranties of merchantability and fitness for a particular use.  " & cName & " County does not warrant the information against deficiencies of any kind.  Under no circumstances will " & cName & " County, or any of its officers or employees be liable for any consequential, incidental, special or exemplary damages even if appraised of the likelihood of such damages occurring.</td>") %>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</body>
</html>
