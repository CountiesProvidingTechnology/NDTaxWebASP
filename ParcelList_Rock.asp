<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Parcel Search Results</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
	Dim cid
	cid = Session("CountyID")
	response.Write("<link rel='stylesheet' href='" & cid & ".css' type='text/css'>")
%>
</head>
<!-- #include file="insDB.asp" -->
<body>
<%
If cName = "Douglas" Then
	Response.Write("<input name='returnHomeButton' type='button' value='Return to Douglas Home Page' onClick=window.location='http://www.co.douglas.mn.us'>")
End If
Dim objCommand, objCommand2, objRS, objRSV, objRSV2, strParcelNo, intPrcl, intRC,  strAddress, Addr, strQueryString, intSet, strTitle1, strTitle2, strTitle3, strTitle4, strTitle5, strName, intPrclNo, strAddr, strNam
strParcelNo = Request.Form("ParcelNo")
Session("PrclNo") = strParcelNo
strAddress = Request.Form("pAddress")
strSAddress = Request.Form("sAddress")
Session("Addr") = strAddress
Session("sAddr") = strSAddress
Addr = Session("Addr")
strName = Request.Form("pName")
Session("Nam") = strName

intSect = Request.Form("sName")
Session("SecNam") = intSect
intTwp = Request.Form("tName")
Session("TwnNam") = intTwp
intRange = Request.Form("rName")
Session("RngNam") = intRange
intRC = Session("intREC")
If strParcelNo <> "" Then
	strParcelNo = replace(strParcelNo, "*", "%")
	'Response.Write(" the second value of : " & strParcelNo)
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL Like '" & strParcelNo & "'"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL Like '" & strParcelNo & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR;"
	strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL Like '" & strParcelNo & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strParcelNo & "' AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR = 2006;"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strParcelNo & "';"
	intSet = 1
	strTitle1 = "Parcel No."
	strTitle2 = "Pay Yr"
	'strTitle3 = "Value/Tax"
	strTitle3 = "E911 Address"
	strTitle4 = "Name"
	strTitle5 = "Description"
End IF
If strAddress <> "" or strSAddress <> "" Then
	strAddress = replace(strAddress, "*", "%")
	strSAdress = replace(strSAdress, "*", "%")
	strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPADR1 Like '" & strAddress & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
	intSet = 2
	strTitle1 = "E911 Address"
	strTitle2 = "Parcel No."
	strTitle3 = "Pay Yr"
	'strTitle4 = "Value/Tax"
	strTitle4 = "Name"
	strTitle5 = "Description"
End IF
If strName <> "" Then
	strName = replace(strName, "*", "%")
	'objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] INNER JOIN NameSearch ON [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = NameSearch.TXPRCL WHERE [NAMESEARCH].TXTNAM Like '" & strName & "'"
	'objCommand.CommandType = 1
	'Set objRS = objCommand.Execute
	'Set objCommand = Nothing
	'strQueryString = "WHERE [NAMESEARCH].TXTNAM Like '" & strName & "'"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM Like '" & strName & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL;"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM Like '" & strName & "' OR [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXANAM Like '" & strName & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR;"
	strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM Like '" & strName & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
	'strQueryString = "WHERE [NAMEQUERY].TXTNAM Like '" & strName & "' ORDER BY [NAMEQUERY].TXPRCL;"
	'strQueryString = "WHERE [NAMESEARCH].TXTNAM Like '" & strName & "' ORDER BY [NAMESEARCH].TXPRCL;"
	'The line below is the good to go previous SQL"""""""""""""""""
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM Like '" & strName & "'"
	intSet = 3
	strTitle1 = "Name"
	strTitle2 = "Parcel No."
	strTitle3 = "Pay Yr"
	'strTitle4 = "Value/Tax"
	strTitle4 = "Description"
End IF
If intTwp <> "" and intRange <> "" Then
	if intSect <> "" Then
	'strQueryString = "WHERE  [STRQuery].TXTOWN = " & intTwp & " AND [STRQuery].TXRANG = " & intRange & " AND [STRQuery].TXSECT = " & intSect & " ORDER BY [STRQuery].TXPRCL, [STRQuery].TXYEAR;"
	strQueryString = "WHERE  [Table 1 - Name/Addr/Desc/Tax/Recap Info ].TXTOWN = " & intTwp & " AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXRANG = " & intRange & " AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXSECT = " & intSect & " ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
	else
	'strQueryString = "WHERE  [STRQuery].TXTOWN = " & intTwp & " AND [STRQuery].TXRANG = " & intRange & " ORDER BY [STRQuery].TXPRCL, [STRQuery].TXYEAR;"
	strQueryString = "WHERE  [Table 1 - Name/Addr/Desc/Tax/Recap Info ].TXTOWN = " & intTwp & " AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXRANG = " & intRange & " ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
	end if
	intSet = 4
	strTitle1 = "Parcel No."
	strTitle2 = "Pay Yr"
	'strTitle3 = "Value/Tax"
	strTitle3 = "E911 Address"
	strTitle4 = "Name"
	strTitle5 = "Description"
End IF

'Check the form and see which field is filled in.
'Then use the correct SQL statement to get the correct RecordSet for the text box filled in.
'*****************************************************************************************
Set objCommand = Server.CreateObject("ADODB.Command")

'Fill in the command properties
objCommand.ActiveConnection = strConnect
if strName <> "" then
	'objCommand.CommandText = "SELECT * FROM [NameQuery] " & strQueryString
	'Line below is the current model ....................................................
	objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
	'objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] INNER JOIN NameSearch ON [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = NameSearch.TXPRCL WHERE [NAMESEARCH].TXTNAM Like '" & strName & "'"
	'objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] INNER JOIN NameSearch ON [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = NameSearch.TXPRCL WHERE [NAMESEARCH].TXTNAM = 'CLARK/ETHEL N';"
	'objCommand.CommandText = "SELECT * FROM NameSearch WHERE [NAMESEARCH].TXTNAM = 'CLARK/ETHEL N';"
	'objCommand.CommandText = "SELECT * FROM NameSearch WHERE [NAMESEARCH].TXTNAM LIKE '" & strName & "';"
	'objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] INNER JOIN NameSearch ON [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = NameSearch.TXPRCL "
else
	if intTwp <> "" and intRange <> "" Then
		objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
	else
		objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
	end if
end if
'objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
'If intSet = 3
'else
objCommand.CommandType = 1
'Response.Write("the query 1 : " & strQueryString )
Set objRS = objCommand.Execute
Set objCommand = Nothing

'end If


Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' ORDER BY [Table 9 - Value Info].PARCEL, [Table 9 - Value Info].YEAR;"

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRSV = objCommand.Execute
Set objCommand = Nothing

'CREATE THE OBJRSV2 FROM TABLE 9 **********************

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' ORDER BY [Table 9 - Value Info].Parcel, [Table 9 - Value Info].Year DESC, [Table 9 - Value Info].RecNum DESC;"

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRSV2 = objCommand.Execute

Set objCommand = Nothing
'objRSV2.close


Response.Write("<table width='1200'><th class='STitle'nowrap>" & strTitle1 & "</th><th class='STitle'>" & strTitle2 & "</th><th class='STitle'>" & strTitle3 & "</th><th class='STitle'>" & strTitle4 & "</th><th class='STitle'>" & strTitle5 & "</th>")

'Response.Write("the session vars are  " & intPrclNo &  cid & Addr & Session("Nam") & Session("SecNam") & Session("TwnNam") & Session("RngNam") & " the end of it ")
'Response.Write("the query : " & strQueryString )
'intPrcl = objRS("TXPRCL")
'Session("intREC") = objRSV2("RecNum")
'Response.Write("value for intSet : " & intSet )
'Response.Write("the query : " & strQueryString )
'Response.Write("the objRSparcel : "  & objRS("TXPRCL"))
While Not objRS.EOF
'Response.Write("the query : " & strQueryString )
	Select Case intSet

	Case 1
		if intPrcl = objRS("TXPRCL") then
			If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then
			else
				Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel_Rock.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=1&yr=" & objRS("TXYEAR") & "'>" & objRS("TXPRCL") & "</a></td>")
				Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
				Response.Write("<td class='rText'>" & objRS("TXPADR1") & "</td>")
				'Response.Write("<td class='rText'>" & objRS("TXFLAG") & "</td>")
				response.Write("<td class='rText'>" & objRS("TXTNAM") & "</td>")
				If  objRS("TXSECT") > 0 Then
					If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				Else
					If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
						Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				End If

				intPrcl = objRS("TXPRCL")
			end If
		else
			Session("intYEAR") = objRS("TXYEAR")
			Session("intPYR") = (objRS("TXYEAR")-1)
			Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0'>" & objRS("TXPRCL") & "</a></td>")
			Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
			Response.Write("<td class='rText'>" & objRS("TXPADR1") & "</td>")
			response.Write("<td class='rText'>" & objRS("TXTNAM") & "</td>")
			If  objRS("TXSECT") > 0 Then
				If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
					Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				else
					Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				end if
			Else
				If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
					Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				else
					Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				end if
			End If
			intPrcl = objRS("TXPRCL")
		End If
	Case 2
		'intPrcl = objRS("TXPRCL")
		'Response.Write(" value for the query string : " &  strQueryString )
		'Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0'>" & objRS("TXPADR") & "</a></td>")
		'Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
		'Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
		'Response.Write("<td class='rText'>" & objRS("TXTNAM") & "</td>")
		'If  objRS("TXSECT") > 0 Then
		'	If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
		'		Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
		'	else
		'		Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
		'	end if
		'Else
		'	If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
		'		Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
		'	else
		'		Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
		'	end if
		'End If
		'objRS.MoveNext
		if intPrcl = objRS("TXPRCL")  then
			'Session("intYEAR") = objRS("TXYEAR")
			If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then
			else
				Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel_Rock.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=1&yr=" & objRS("TXYEAR") & "'>" & objRS("TXPADR1") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
				Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
				Response.Write("<td class='rText'>" & objRS("TXTNAM") & "</td>")
				If  objRS("TXSECT") > 0 Then
					If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				Else
					If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
						Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				End If
				intPrcl = objRS("TXPRCL")
			End If
		else
			Session("intYEAR") = objRS("TXYEAR")
			Session("intPYR") = (objRS("TXYEAR")-1)
			Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0'>" & objRS("TXPADR1") & "</a></td>")
			Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
			Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
			Response.Write("<td class='rText'>" & objRS("TXTNAM") & "</td>")
			If  objRS("TXSECT") > 0 Then
				If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
					Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				else
					Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				end if
			Else
				If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
					Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				else
					Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				end if
			End If
			intPrcl = objRS("TXPRCL")
		End If

	Case 3
		'intPrcl = objRS("TXPRCL")
		'Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0'>" & objRS("TXTNAM") & "</a></td>")
		'Response.Write("<td class='rText'>" & objRS("TXPRCL") & "</td>")
		'Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
		'If  objRS("TXSECT") > 0 Then
		'	If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
		'		Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
		'	else
		'		Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
		'	end if
		'Else
		'	If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
		'		Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
		'	else
		'		Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
		'	end if
		'End If
		'objRS.MoveNext
		if intPrcl = objRS("TXPRCL") then
				'Session("intYEAR") = objRS("TXYEAR")
			If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then
			else
				Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel_Rock.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=1&yr=" & objRS("TXYEAR") & "'>" & objRS("TXTNAM") & objRS("TXANAM") &" </a></td>")
				'Response.Write("<td class='rText'>" & objRS("TXTNAM") &  "</td>")
				Response.Write("<td class='rText'>" & objRS("TXPRCL") & "</td>")
				Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
				If  objRS("TXSECT") > 0 Then
					If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				Else
					If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
						Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				End If
				intPrcl = objRS("TXPRCL")
			End If
		else
				Session("intYEAR") = objRS("TXYEAR")
				Session("intPYR") = (objRS("TXYEAR")-1)
				Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0'>" & objRS("TXTNAM") & objRS("TXANAM") & "</a></td>")
				Response.Write("<td class='rText'>" & objRS("TXPRCL") & "</td>")
				Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
				If  objRS("TXSECT") > 0 Then
					If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				Else
					If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
						Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				End If
				intPrcl = objRS("TXPRCL")
		End If
	Case 4
			'Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0'>" & objRS("TXPRCL") & "</a></td>")
			'Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
			'Response.Write("<td class='rText'>" & objRS("TXPADR1") & "</td>")
			'response.Write("<td class='rText'>" & objRS("TXTNAM") & "</td>")
			'If  objRS("TXSECT") > 0 Then
			'		If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
			'			Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			'		else
			'			Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			'		end if
			'Else
			'		If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
			'			Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			'		else
			'			Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			'		end if
			'End If

			if intPrcl = objRS("TXPRCL") then
					'Session("intYEAR") = objRS("TXYEAR")
				If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then
				else
						Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel_Rock.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=1&yr=" & objRS("TXYEAR") & "'>" & objRS("TXPRCL") & "</a></td>")
						Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
						Response.Write("<td class='rText'>" & objRS("TXPADR1") & "</td>")
						response.Write("<td class='rText'>" & objRS("TXTNAM") & "</td>")
						If  objRS("TXSECT") > 0 Then
							If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
								Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
							else
								Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
							end if
						Else
							If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
								Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
							else
								Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
							end if
						End If
						intPrcl = objRS("TXPRCL")
				End If
			else
					Session("intYEAR") = objRS("TXYEAR")
					Session("intPYR") = (objRS("TXYEAR")-1)
					Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0'>" & objRS("TXPRCL") & "</a></td>")
					Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
					Response.Write("<td class='rText'>" & objRS("TXPADR1") & "</td>")
					response.Write("<td class='rText'>" & objRS("TXTNAM") & "</td>")
					If  objRS("TXSECT") > 0 Then
						If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
							Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
						else
							Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
						end if
					Else
						If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
							Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
						else
							Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
						end if
					End If
					intPrcl = objRS("TXPRCL")
			End If

		'Response.Write("<td class='rText'>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") & " " & objRS("TXDSC1") & "</td></tr>")
	End Select
	objRS.MoveNext
	'If objRS.EOF  then
	'Response.Write("End Of File")
	'end if
Wend

objRS.Close
Set objRS = Nothing

'Session("intREC") = objRSV2("RecNum")
%>

</table>
</body>
</html>
