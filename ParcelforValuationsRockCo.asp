<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Parcel Search Results</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
	Dim cid, cName, sesTotal, sesParcel, sesRecNum, sesTotalRec
	cid = Session("CountyID")
	cName = Session("County Name")
	sesTotal = Session("amtTotal")
	sesParcel = Session("strPID")
	sesRecNum = Session("counter")
	sesTotalRec = Session("intTotalRec")
	RecNumCase = sesRecNum
	'Response.Write("sesRecNum#1= ") & sesRecNum
	'Response.Write("RecNumCase#1= ") & RecNumCase
	'Response.Write("intTotalRec#1= ") & intTotalRec
	'Response.Write("sesTotalRec#1= ") & sesTotalRec

	If sesRecNum > sesTotalRec Then
			'count = 1
			'Session("count") = count
			'sesRecNum = Session("count")
			'sesRecNum = 1
			RecNumCase = sesRecNum
	'		Response.Write("sesRecNum#11= ") & sesRecNum
	'		Response.Write("RecNumCase#11= ") & RecNumCase
	End If


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


Function calcZip(strData)
	If objRS(strData) = "00000" Then
		strData = ""
	End If
	calcZip = strData
end Function

Dim objCommand, objRS, strQueryString, strPID, strTID, objRS3, objRS5, objRS94, objRS95, objRS93, intNumRecords, intNumRecords2, intI3, intI4, intI5, intRec3, strParcel, RecNumCase
Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strParcel = strPID
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
'Response.Write("strQueryString = ") & strQueryString
strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strQueryString & "'"
'Response.Write("strQueryString = ") & strQueryString
'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
objCommand.CommandType = 1

Set objRS = objCommand.Execute

Set objCommand = Nothing

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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'to create recordset RS94 for Valuations page YEAR 2004
Set objCommand = Server.CreateObject("ADODB.Command")
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
'strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 1 AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND  [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
'strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 0 AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"


'Fill in the command properties
objCommand.ActiveConnection = strConnect
'objRS94.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1
'Set objRS94 = Server.CreateObject("ADODB.Recordset")
Set objRS94 = objCommand.Execute
'objRS94.ActiveConnection = strConnect

'objRS94.CursorType = adOpenStatic
'objRS94.Open "SELECT * FROM [Table 9 - Value Info] WHERE ([Table 9 - Value Info].YEAR = 2004 AND  [Table 9 - Value Info].Parcel = '" & strQueryString & "')" , strConnect, addOpenStatic, adLockReadOnly, adCmdTable
'intNumRecords = objRS94.RecordCount
Set objCommand = Nothing

intI4 = 0
Do While Not objRS94.EOF
	If objRS94("RecNum") >= 0   Then
		If objRS94("Deleted") <> ""    Then
		Else
			intI4 = intI4 + 1
		End If
	End If
	objRS94.MoveNext
Loop

objRS94.Close

'to create recordset RS94 for Valuations page YEAR 2004
Set objCommand = Server.CreateObject("ADODB.Command")
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
If intI4 = 1  Then
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 1 AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Else
	'strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND  [Table 9 - Value Info].Parcel = '" & strQueryString & "')"'

	RecNumCase = sesRecNum

	'if sesRecNum < 1 Then
	'RecNumCase = 0
	'end if
	'Response.write("sesRecNum= ") & sesRecNum'



	'Response.write("sesRecNum= ") & sesRecNum
	'Response.write("RecNumCase= ") & RecNumCase
	'Response.write("sesTotalRec= ") & intTotalRec


	'strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 0  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"'


Select Case sesRecNum

Case  0
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 0  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  1
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 0  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  2
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 1  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  3
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 2  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  4
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 3  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  5
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 4  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  6
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 5  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  7
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 6  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  8
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 7  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  9
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 8  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  10
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2004 AND [Table 9 - Value Info].RecNum = 9  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
End Select

End If

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1
Set objRS94 = objCommand.Execute
'intNumRecords = objRS94.RecordCount
Set objCommand = Nothing


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'to create recordset RS95 for Valuations page YEAR 2005
Set objCommand = Server.CreateObject("ADODB.Command")
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
'strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 1 AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRS95 = objCommand.Execute
'intNumRecords2 = objRS95.RecordCount
Set objCommand = Nothing

intI5 = 0
Do While Not objRS95.EOF
	If objRS95("RecNum") >= 0   Then
		If objRS95("Deleted") <> ""    Then

		Else
			intI5 = intI5 + 1
		End If
	End If
	objRS95.MoveNext
Loop
Session("intTotalRec") = intI5
'Response.Write("intTotalRec= ") & intI5
If sesRecNum > intI5 Then
sesRecNum = 1
End if

objRS95.Close

'to create recordset RS95 for Valuations page YEAR 2005
Set objCommand = Server.CreateObject("ADODB.Command")
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
If intI5 = 1 Then
strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 1 AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Else
'strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 0 AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
'strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"

Select Case sesRecNum

Case  0
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 0  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  1
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 0  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  2
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 1  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  3
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 2  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  4
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 3  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  5
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 4  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  6
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 5  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  7
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 6  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  8
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 7  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  9
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 8  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  10
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2005 AND [Table 9 - Value Info].RecNum = 9  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
End Select

End If

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRS95 = objCommand.Execute
'intNumRecords2 = objRS95.RecordCount
Set objCommand = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'to create recordset RS93 for Valuations page YEAR 2003
Set objCommand = Server.CreateObject("ADODB.Command")
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")

'strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 1 AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRS93 = objCommand.Execute
Set objCommand = Nothing

intI3 = 0
Do While Not objRS93.EOF
	If objRS93("RecNum") >= 0   Then

		If objRS93("Deleted") <> ""    Then

		Else
		intI3 = intI3 + 1
		End If
	End If
	objRS93.MoveNext
Loop

objRS93.Close

'to create recordset RS93 for Valuations page YEAR 2003
Set objCommand = Server.CreateObject("ADODB.Command")
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
If intI3 = 1 Then
strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 1 AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Else
'intRec3 = 1
'strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 0 AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"

Select Case sesRecNum

Case  0
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 0  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  1
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 0  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  2
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 1  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  3
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 2  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  4
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 3  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  5
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 4  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  6
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 5  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  7
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 6  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  8
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 7  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  9
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 8  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
Case  10
	strQueryString = "WHERE ([Table 9 - Value Info].YEAR = 2003 AND [Table 9 - Value Info].RecNum = 9  AND [Table 9 - Value Info].Parcel = '" & strQueryString & "')"
End Select

End If
'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRS93 = objCommand.Execute
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
			If cName = "Rock" Then
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5' class='uLink'>Valuations</a>")
			else
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
			end if
			Case 1
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='hLink'>Tax Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>|")
			If cName = "Rock" Then
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5' class='uLink'>Valuations</a>")
			else
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
			end if
			Case 2
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='hLink'>Current Receipts</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>|")
			If cName = "Rock" Then
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5' class='uLink'>Valuations</a>")
			else
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
			end if
			Case 3
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='hLink'>Unpaid Tax</a>|")
			If cName = "Rock" Then
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5' class='uLink'>Valuations</a>")
			else
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
			end if
			Case 4
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>|")
			If cName = "Rock" Then
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='hLink'>History</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5' class='uLink'>Valuations</a>")
			else
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='hLink'>History</a>")
			end if
			Case 5
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>|")
			If cName = "Rock" Then
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5' class='hLink'>Valuations</a>")
			else
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
			end if
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
			If cName = "Douglas" Then
				Response.Write("<a href='http://www.co.douglas.mn.us' class='tlink'>Douglas County Home Page</a>")
			End If
			If cName = "Rock" Then
				Response.Write("<a href='http://www.co.rock.mn.us' class='tlink'>Rock County Home Page</a>")
			End If
			If cName = "Norman" Then
				Response.Write("<a href='http://www.co.norman.mn.us/' class='tlink'>Norman County Home Page</a>")
			End If
			If cName = "Pope" Then
				Response.Write("<a href='#' class='tlink'></a>")
			End If
			If cName = "Kandiyohi" Then
				Response.Write("<a href='http://www.co.kandiyohi.mn.us/' class='tlink'>Kandiyohi County Home Page</a>")
			End If
			If cName = "Renville" Then
				Response.Write("<a href='http://www.co.renville.mn.us/' class='tlink'>Renville County Home Page</a>")
			End If
			If cName = "Stevens" Then
				Response.Write("<a href='http://www.co.stevens.mn.us/' class='tlink'>Stevens County Home Page</a>")
			End If
%>
			</td>
	</tr>
</table>
<%
objRS.Close
Set objRS = Nothing
objRS95.Close
Set objRS95 = Nothing
objRS94.Close
Set objRS94 = Nothing
objRS93.Close
Set objRS93 = Nothing




%>
<form action="searchinputreturn.asp" method="post">
<center>
<%
'Session("counter") = 1
'sesRecNum = 1
%>
<input name="decisionButton" type="submit" value="Another Search">&nbsp;&nbsp;&nbsp;
</center>
</form>

</body>
</html>
