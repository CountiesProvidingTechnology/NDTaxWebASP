<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>County Parcel Search Page</title>
<%
	Dim varcid, cName
	'cid = Session("CountyID")
	varcid = request.QueryString("cid")
	'Get the county name based on the county number
'Get the county name based on the county number
		Select Case varcid
				Case 21
				'Douglas County
				cName = "Douglas"
				Case 26
				'Grant County
				cName = "Grant"
				Case 67
				'Rock County
				cName = "Rock"
				Case 53
				'Nobles County
				cName = "Nobles"
				Case 54
				'Norman County
				cName = "Norman"
				Case 64
				'Redwood County
				cName = "Redwood"
				Case 65
				'Renville County
				cName = "Renville"
				Case 74
				'Steele County
				cName = "Steele"
				Case 75
				'Stevens County
				cName = "Stevens"
				Case 76
				'Swift County
				cName = "Swift"
				Case 77
				'Todd County
				cName = "Todd"
				Case 87
				'Yellow Medicine County
				cName = "Yellow Medicine"
				Case 61
				'Pope County
				cName = "Pope"
				Case 34
				'Kandiyohi County
				cName = "Kandiyohi"
				Case 47
				'Meeker County
				cName = "Meeker"
				Case 48
				'Mille Lacs County
				cName = "Mille Lacs"
				Case 42
				'Lyon County
				cName = "Lyon"
				Case 45
				'Marshall County
				cName = "Marshall"
				Case 51
				'Murray County
				cName = "Murray"
				Case 41
				'Lincoln County
				cName = "Lincoln"
				Case 78
				'Traverse County
				cName = "Traverse"
				Case 84
				'Wilkin County
				cName = "Wilkin"
				Case Else
				'Kandiyohi County
				cName = "Kandiyohi"
		End Select
	'cName = Session("County Name")
	response.Write("<link rel='stylesheet' href='" & varcid & ".css' type='text/css'>")
%>
</head>
<!-- #include file="insDB.asp" -->
<body text="#000000" bgcolor="#FFFFFF">
<%
	'If cid = "" Then
	'	response.Redirect("tax.asp")
	'End If

	'Dim strResponse, strSearchAgain
	'strResponse = Request.Form("decision")
	'strSearchAgain = request.QueryString("Again")

	'If strResponse = "iDecline" Then
	'	response.Redirect("Decline.asp")
	'Else
	'	If strResponse = "" AND strSearchAgain = "" Then
	'		response.Redirect("Disclaimer.asp")
	'	End If
	'End If


'Empty out the session variables for the searches . This will allow the user to have a fresh start
	' on each new click to do a new search from the Parcel.asp page.'
	intZero = 0
	'Session("PrclNo") = ""
	'Session("Addr") = ""
	'Session("Nam") = ""

	'Session("SecNam") = ""

	'Session("TwnNam") = ""

	'Session("RngNam") = ""
	'Response.Write("value for session name " & Session("Nam") & " value for session address " & Session("Addr") & " value for session parcel " & Session("PrclNo") )
	'end the added code here for the clearing of the session variables when coming from the parcel.asp page.'



	Dim objCommand, objRS, strQueryString, strPID, strTID, objRS3, objRS5

	Set objCommand = Server.CreateObject("ADODB.Command")

	'strPID = request.QueryString("pid")
	'strTID = request.QueryString("tid")
	'strQueryString = request.QueryString("pid")
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strQueryString & "'"

	'Fill in the command properties
	objCommand.ActiveConnection = strConnect
	objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] "
	objCommand.CommandType = 1

	Set objRS = objCommand.Execute

	Set objCommand = Nothing


	If cName = "Douglas" Then
		Response.Write("<img src='Images/depart23.gif'>")
	End If
%>

<form action="ParcelList_value.asp" method="post">
<input type="hidden" name="cid" value=<%=varcid%> >
<table border="0" cellspacing="0" cellpadding="0">
	<tr valign="top">
		<td width="100"></td>
		<td width="550">
		<%
				If varcid = 61 Then        ' Just for Pope county this is the order of the search boxes that they wanted. LEM 03/02/09
			%>
			<!-- #include file="SearchByName.asp" -->
				<br>
				<br>
			<!-- #include file="SearchByParcel.asp" -->
				<br>
				<br>
			<!-- #include file="SearchByE911_Pope.asp" -->
				<br>
				<br>
			<!-- #include file="SearchBySectTwpRange.asp" -->
				
			<%
				End If
			%>
			<%
				If varcid = 54 or varcid = 34 or varcid = 75 or varcid = 67 or varcid = 48 or varcid = 87 or varcid = 65 or varcid = 53 or varcid = 42 Then
			%>
			<!-- #include file="SearchByName.asp" -->
			<%
				End If
			%>
				<br>
				<br>
			<%
				If varcid = 54 or varcid = 34 or varcid = 21 or varcid = 48 or varcid = 26 or varcid = 65 or varcid = 67 or varcid = 75 or varcid = 87 or varcid = 76 or varcid = 53 or varcid = 47 or varcid = 74 or varcid = 42 or varcid = 84 Then
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
					<td width="450" class="rText">The use of a wild card is permitted.  The wild card for this search is '*'.  <br>For example:<br><b>03*</b> in the Parcel Search field would result in a list of all parcels that begin with 03.<br>And<br>
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
