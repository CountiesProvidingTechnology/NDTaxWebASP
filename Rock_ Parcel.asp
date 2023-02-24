<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Parcel Search Results</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
	Dim cid, cName, sesTotal
	cid = Session("CountyID")
	cName = Session("County Name")
	sesTotal = Session("amtTotal")
	response.Write("<link rel='stylesheet' href='" & cid & ".css' type='text/css'>")
%>
</head>
<!-- #include file="insDB.asp" -->


<body>
<%
Function printTaxRecord()
	Dim recordNumber
	For recordNumber=1 to 5
		If objRS3("RCBTDT" & recordNumber) = 0 Then
			Response.Write("<tr valign='top'>")
				If recordNumber = 1 Then
					Response.Write("<td class='rText' align='left' rowspan='7'>No Tax Receipt Information</td>")
				Else
					'Response.Write("<td class='rText' align='left' rowspan='6'>&nbsp;</td>")'
				End If
				'Response.Write("<td width='60' class='rText' align='right'>&nbsp;</td>")
				'Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				'Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				'Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				'Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				'Response.Write("<td width='80' class='rText' align='right'> top " & recordnumber & subrecordnumber & tempsubrecordnumber & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")'
			if recordnumber < 5 then
			If objRS3("RCBTDT" & (recordNumber + 1)) = 0 Then
			'Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")'
			else
				'Response.Write("<td width='80' class='rText' align='right'>" & recordnumber & subrecordnumber & "</td>")
			end if
			end if
			'Response.Write("</tr>")
			recordNumber = 5
		Else
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='260' class='rText' align='left' rowspan='6'>")
				'Response.Write("<b>" & calcDate(("RCBTDT" & recordNumber), (recordNumber)) & "</b><br>")'
				Response.Write("<b>" & calcDate(("RCBTDT"), recordnumber) & "</b><br>")
				Response.Write("<b>Batch # " & objRS3(("RCBAT#" & recordNumber)) & "</b><br>")
				Response.Write("<b>Paid</b> by " & objRS3(("RCPDBY" & recordNumber)) & "<br>")
				Response.Write("<b>Validation #</b> " & objRS3(("RCVAL#" & recordNumber)))
				Response.Write("</td>")
				'Response.Write("<td width='60' class='rText' align='right'>" & objRS3(("RCTYP" & recordNumber & "1")) & "</td>")'
				Response.Write("<td width='60' class='rText' align='right'>" & objRS3(("RCTYP"  & "1" & recordNumber)) & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCAMT" & recordNumber & "1")), 2) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCAMT"  & "1" & recordNumber)), 2) & "</td>")
				sesTotal = objRS3("RCAMT" & "1" & recordNumber)
				'Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAR" & recordNumber & "1")) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAR"  & "1" & recordNumber)) & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAC" & recordNumber & "1")) & "</td>")'
				'Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAC" & recordNumber & "1" & recordNumber)) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAC" & "1" & recordNumber)) & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCSAA" & recordNumber & "1")), 2) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCSAA" & "1" & recordNumber)), 2) & "</td>")
				sesTotal = sesTotal + objRS3("RCSAA" & "1" & recordNumber)
				Response.Write("</tr>")
			printTaxValues(recordNumber)
		End If
	Next
End Function

'Function printTaxValues()'
Function printTaxValues(recordNumber)
	'Dim recordValue'
	Dim subrecordNumber, tempsubrecordnumber
	'For recordValue=2 to 6'
	For subrecordNumber=2 to 6
		'If objRS3(("RCAMT" & recordValue & "1")) > 0 Then'
		'Response.Write("the entry of printTaxValues 2-6 " & subrecordNumber )'
		If (objRS3(("RCSAA" & subrecordNumber & recordNumber )) > 0) or (objRS3(("RCAMT" & subrecordNumber & recordNumber))>0) Then
				Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText' align='right'>"  & objRS3(("RCTYP" & subrecordNumber & recordNumber)) & "</td>")
				If objRS3(("RCAMT" & subrecordNumber & recordNumber )) > 0 Then
					Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCAMT" & subrecordNumber & recordNumber)), 2) & "</td>")
					sesTotal = sesTotal + objRS3("RCAMT" & subrecordNumber & recordNumber )
				Else
					Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				End If
				If objRS3("RCSAA" & subrecordNumber & recordNumber) > 0 Then
				Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAR" & subrecordNumber & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAC" & subrecordNumber & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCSAA" & subrecordNumber & recordNumber)), 2) & "</td>")
				Response.Write("</tr>")
				else
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				Response.Write("</tr>")

				End If
					sesTotal = sesTotal + objRS3("RCSAA" & subrecordNumber & recordNumber)
		Else
				Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				'Response.Write("<td width='80' class='rText'>" & recordnumber & "</td>")
				Response.Write("</tr>")
				'Response.Write("entry of printTaxValues 2-6 " & subrecordNumber )
		End If

		If (objRS3(("RCSAA"& subrecordNumber & recordNumber )) = 0) and (objRS3(("RCAMT" & subrecordNumber & recordNumber))=0) Then
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				'Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("</tr>")

			Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				'Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("</tr>")

				Response.Write("<tr valign='top'>")
					Response.Write("<td width='60' class='rText'>&nbsp;</td>")
					Response.Write("<td width='80' class='rText'>&nbsp;</td>")
					Response.Write("<td width='80' class='rText'>&nbsp;</td>")
					'Response.Write("<td width='80' class='rText'>"& subrecordnumber & recordNumber & "</td>")
				If (objRS3(("RCSAA2" & recordNumber )) = 0) or (objRS3(("RCSAA3" & recordNumber )) = 0) Then
				else
					Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				End if
					Response.Write("<td width='80' class='rText'>Total</td>")
					Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber((sesTotal), 2) & "</td>")
				Response.Write("</tr>")
					Response.Write("<tr valign='top'>")
						Response.Write("<td width='60' class='rText'>&nbsp;</td>")
						Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						'Response.Write("<td width='80' class='rText'>&nbsp;</td>")'

					If  (objRS3(("RCAMT2" & recordNumber )) = 0) and (objRS3(("RCSAA3" & recordNumber )) = 0) Then
							'Response.Write("<td width='80' class='rText'>RCAMT2=0" & subrecordnumber & recordnumber & "</td>")
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						if subrecordnumber > 2 then
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						end if
					else
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
							'Response.Write("<td width='80' class='rText'>RCAMT2<>0" & subrecordnumber & recordnumber & "</td>")
						if (objRS3(("RCAMT2" & recordNumber + 1 )) = 0) or (objRS3(("RCSAA3" & recordNumber + 1 )) = 0) Then
							'Response.Write("<td width='80' class='rText'>RCSAA3+1" & subrecordnumber & recordnumber & "</td>")
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						else
							'Response.Write("<td width='80' class='rText'>RCSAA3+1" & subrecordnumber & recordnumber & "</td>")
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						end if
					end if
						Response.Write("</tr>")
				tempsubrecordnumber = subrecordnumber
				subrecordNumber = 7
		End if
		Next
	'end if
	End Function





'code below is not used and is original but commented out '
'		If objRS3(("RCSAA" & recordValue & "1")) > 0 Then
'			Response.Write("<tr valign='top'>")
'			Response.Write("<td width='60' class='rText' align='right'>" & objRS3(("RCTYP" & recordValue & "1")) & "</td>")
'			If objRS3(("RCAMT" & recordValue & "1")) > 0 Then
'				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCAMT" & recordValue & "1")), 2) & "</td>")
'			Else
'				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
'			End If
'			Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAR" & recordValue & "1")) & "</td>")
'			Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAC" & recordValue & "1")) & "</td>")
'			Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCSAA" & recordValue & "1")), 2) & "</td>")
'			Response.Write("</tr>")
'		Else
'			Response.Write("<tr valign='top'>")
'			Response.Write("<td width='60' class='rText'>&nbsp;</td>")
'			Response.Write("<td width='80' class='rText'>&nbsp;</td>")
'			Response.Write("<td width='80' class='rText'>&nbsp;</td>")
'			Response.Write("<td width='80' class='rText'>&nbsp;</td>")
'			Response.Write("<td width='80' class='rText'>&nbsp;</td>")
'			Response.Write("</tr>")
'		End If
'	Next
'End Function

Function calcZero(strData, intPlaces)
	If objRS(strData) <> "" Then
		strData = (FormatNumber(objRS(strData), intPlaces))
	Else
		strData = "0.00"
	End If
	calcZero = strData
end Function

Function calcDate(strData, intrecordnumber)
	Dim strYear, strMonth, strDay, intLength
	strData = objRS3("RCBTDT" & intrecordnumber)
	intLength = Len(strData)
	strYear = Right(strData, 4)
	If intLength = 7 Then
		strMonth = Left(strData, 1)
		strDay = Mid(strData, 2, 2)
	Else
		strMonth = Left(strData, 2)
		strDay = Mid(strData, 3, 2)
	End If
	calcDate = strMonth + "/" + strDay + "/" + strYear
end Function

Function calcDateRS4(strData, intrecordnumber)
	Dim strYear, strMonth, strDay, intLength
	strData = objRS4("RCBTDT" & intrecordnumber)
	intLength = Len(strData)
	strYear = Right(strData, 4)
	If intLength = 7 Then
		strMonth = Left(strData, 1)
		strDay = Mid(strData, 2, 2)
	Else
		strMonth = Left(strData, 2)
		strDay = Mid(strData, 3, 2)
	End If
	calcDateRS4 = strMonth + "/" + strDay + "/" + strYear
end Function


Function calcZip(strData)
	If objRS(strData) = "00000" Then
		strData = ""
	End If
	calcZip = strData
end Function

Dim objCommand, objRS, strQueryString, strPID, strTID, objRS3, objRS5, objRScount, intnumberQ

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strQueryString & "'"
'Response.Write("The value of the Query String : " & strQueryString )

'Fill in the command properties
'Response.Write( "the strConnect is " & strConnect   )
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
objCommand.CommandType = 1

Set objRS = objCommand.Execute

Set objCommand = Nothing




'Dim objCommand, objRScount, intnumberQ

'Set objCommand = Server.CreateObject("ADODB.Command")


'Fill in the command properties'
'Response.Write( "the strConnectCNT is " & strConnectCNT   )
'intnumberQ = 1
'response.Write(" the cid is " & cid  )
'response.Write("the value of intnumberQ : " & intnumberQ )
'intnumberQ = 1
'strQueryString = "WHERE [Counter" & cid & "].count ='" & intnumberQ & "' "
'Response.Write("the value of the query string :" & strQueryString )
'objCommand.ActiveConnection = strConnectCNT
'strQueryString = "SELECT * FROM [Counter" & cid & "] WHERE [Counter" & cid & "].count ='1'"
'objCommand.CommandText = "SELECT * FROM [Counter" & cid & "] WHERE [Counter" & cid & "].count ='1'"
'Response.Write(" the value of the SELECT part of the Query String : " & strQueryString )
'objCommand.CommandType = 1
'Set objRScount = objCommand.Execute

'Set objCommand = Nothing
'Response.Write("Your session started at : " & Session("Start") )

'Response.Write("there have been " & Session("VisitorID") & " total visits to this site")

'Response.Write("value of COUNT : ")
'Response.Write(  objRScount("COUNT") )
'Response.Write(  objRScount("KEYID") )
'Response.Write(  objRScount("TOTCOUNT") )
'number = objRScount("TOTCOUNT")
'number = objRScount("count")
'Response.Write("the value of number the first time is : " & number )
'objRScount("TOTCOUNT")  = (objRScount("TOTCOUNT")) + 1
'Response.Write("the value of number the second time is : " & number )
'objRScount.update
'number = objRScount("COUNT")
'objRScount.Close
'Set objRScount = Nothing









Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 3 - Part1 of RCPT Sets(1-5)].TXPRCL = '" & strQueryString & "'"

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 3 - Part1 of RCPT Sets(1-5)] " & strQueryString
objCommand.CommandType = 1

Set objRS3 = objCommand.Execute
Set objCommand = Nothing

Set objCommand = Nothing

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 5 - Unpaid Taxes/Truth in Taxation Info].TXPRCL = '" & strQueryString & "'"

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 5 - Unpaid Taxes/Truth in Taxation Info] " & strQueryString
objCommand.CommandType = 1

Set objRS5 = objCommand.Execute
Set objCommand = Nothing
%>

<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="650" align="right" colspan="2">Parcel Number: <b><%= objRS("TXPRCL") %></b></td>
	</tr>
	<tr valign="top">
		<td width="650" align="right" colspan="2">Tax Year: <b><%= objRS("TXYEAR") %></b></td>
	</tr>
	<tr>
		<td width="10"></td>
		<td width="650" align="left">
		<%
			Select Case strTID
			Case 0
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='hLink'>General Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				'If cName = "Douglas" Then
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|   |")
				'	'response.Write("<a href='javascript:zoomto_selectfeature('strPIN');' class='uLink'> View Maps </a>")
				'	response.Write("<a href='http://morris.state.mn.us/dcmap/' class='ulink'>View Maps </a>")
				'else
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				'end if

			Case 1
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='hLink'>Tax Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				'If cName = "Douglas" Then
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|   |")
				'	response.Write("<a href='http://morris.state.mn.us/dcmap/' class='ulink'>View Maps </a>")
				'	'response.Write("<a href='javascript:zoomto_selectfeature('strPIN');' class='uLink'> View Maps </a>")
				'else
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				'end if
			Case 2
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='hLink'>Current Receipts</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
			'If cName = "Douglas" Then
			'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|   |")
			'	response.Write("<a href='http://morris.state.mn.us/dcmap/' class='ulink'>View Maps </a>")
			'else
			'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
			'end if
			Case 3
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='hLink'>Unpaid Tax</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				'If cName = "Douglas" Then
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|   |")
				'	response.Write("<a href='http://morris.state.mn.us/dcmap/' class='ulink'>View Maps </a>")
				'else
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				'end if
			Case 4
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				'If cName = "Douglas" Then
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|   |")
				'	response.Write("<a href='http://morris.state.mn.us/dcmap/' class='ulink'>View Maps </a>")
				'else
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				'end if
			Case 5
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				'If cName = "Douglas" Then
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|   |")
				'	response.Write("<a href='http://morris.state.mn.us/dcmap/' class='ulink'>View Maps </a>")
				'else
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				'end if
			End Select
		%>
		</td>
	</tr>
	<tr>
		<td width="650" bgcolor="#000000" height="1" colspan="2"></td>
	</tr>
	<tr valign="top">
		<td height="15"></td>
	</tr>
	<tr valign="top">
		<td width="10"></td>
		<td width="650">
		<%
			Select Case strTID
			Case 0
		%>
		<!-- #include file="GeneralTax.asp" -->
		<%
			Case 1
		%>
		<!-- #include file="TaxInformation.asp" -->
		<%
			Case 2
		%>
		<!-- #include file="CurrentReceipts.asp" -->
		<%
			Case 3
		%>
		<!-- #include file="UnpaidTax.asp" -->
		<%
			Case 4
		%>
		<!-- #include file="History.asp" -->
		<%
			Case 5
		%>
		<!-- #include file="Valuations.asp" -->
		<%
			End Select
		%>
		</td>
	</tr>
	<tr>
		<td height="15" colspan="2"></td>
	</tr>
	<tr valign="top">
		<td width="10"></td>
		<td width="650" align="right" class="STitle2">
<%
If cName = "Rock" then
	Response.Write("<a href='/tax/ParcelListReturn_Rock.asp' class='tlink'>Back to ParcelList    |</a>&nbsp;&nbsp;&nbsp;&nbsp")
Else
	Response.Write("<a href='/tax/ParcelListReturn.asp' class='tlink'>Back to ParcelList    |</a>&nbsp;&nbsp;&nbsp;&nbsp")
End if
			'If cName = "Douglas" Then
			'	Response.Write("<a href='http://www.co.douglas.mn.us' class='tlink'>Douglas County Home Page</a>")
			'End If
			'If cName = "Rock" Then
			'	Response.Write("<a href='http://www.co.rock.mn.us' class='tlink'>Rock County Home Page</a>")
			'End If
			'If cName = "Norman" Then
			'	Response.Write("<a href='http://www.co.norman.mn.us/' class='tlink'>Norman County Home Page</a>")
			'End If
			'If cName = "Pope" Then
			'	Response.Write("<a href='#' class='tlink'></a>")
			'End If
			'If cName = "Kandiyohi" Then
			'	Response.Write("<a href='http://www.co.kandiyohi.mn.us/' class='tlink'>Kandiyohi County Home Page</a>")
			'End If
			'If cName = "Renville" Then
			'	Response.Write("<a href='http://www.co.renville.mn.us/' class='tlink'>Renville County Home Page</a>")
			'End If
			'If cName = "Stevens" Then
			'	Response.Write("<a href='http://www.co.stevens.mn.us/' class='tlink'>Stevens County Home Page</a>")
			'End If
%>
			</td>
	</tr>
</table>
<%
Response.Write("Your session county at : " & cid )
'Response.Write("Your session started at : " & Session("Start") )
'
'Response.Write("there have been " & Session("VisitorID") & " total visits to this site")
'
'Response.Write("value of COUNT : ")
'Response.Write(  objRScount("COUNT") )
'Response.Write(  objRScount("KEYID") )
'Response.Write(  objRScount("TOTCOUNT") )
'number = objRScount("TOTCOUNT")
'number = objRScount("count")
'Response.Write("the value of number the first time is : " & number )
'objRScount("TOTCOUNT")  = (objRScount("TOTCOUNT")) + 1
'Response.Write("the value of number the second time is : " & number )
'objRScount.update
'number = objRScount("COUNT")




objRS.Close
Set objRS = Nothing

'objRScount.Close
'Set objRScount = Nothing

%>
<form action="searchinputreturn.asp" method="post">
<center>


<input name="decisionButton" type="submit" value="Another Search">&nbsp;&nbsp;&nbsp;
</center>
</form>

</body>
</html>
