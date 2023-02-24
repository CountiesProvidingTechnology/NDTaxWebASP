<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Parcel Search Results</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
	Dim cid, cName, sesTotal
	strPID = request.QueryString("pid")
	strYEAR = request.QueryString("YEAR")
	strCID = request.QueryString("cid")
	Session("CountyID") = strCID
	cid = Session("CountyID")
	if cid = 21 then
	cName = "Douglas"
	end if

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

'The next printTaxRecord Function will use the next table TABLE 4 that holds tax receipts .

Function printTaxRecord4()
	Dim recordNumber
	For recordNumber=1 to 5
		If objRS4("RCBTDT" & recordNumber) = 0 Then
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
			If objRS4("RCBTDT" & (recordNumber + 1)) = 0 Then
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
				Response.Write("<b>Batch # " & objRS4(("RCBAT#" & recordNumber)) & "</b><br>")
				Response.Write("<b>Paid</b> by " & objRS4(("RCPDBY" & recordNumber)) & "<br>")
				Response.Write("<b>Validation #</b> " & objRS4(("RCVAL#" & recordNumber)))
				Response.Write("</td>")
				'Response.Write("<td width='60' class='rText' align='right'>" & objRS4(("RCTYP" & recordNumber & "1")) & "</td>")'
				Response.Write("<td width='60' class='rText' align='right'>" & objRS4(("RCTYP"  & "1" & recordNumber)) & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS4(("RCAMT" & recordNumber & "1")), 2) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS4(("RCAMT"  & "1" & recordNumber)), 2) & "</td>")
				sesTotal = objRS4("RCAMT" & "1" & recordNumber)
				'Response.Write("<td width='80' class='rText' align='right'>" & objRS4(("RCSAR" & recordNumber & "1")) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & objRS4(("RCSAR"  & "1" & recordNumber)) & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>" & objRS4(("RCSAC" & recordNumber & "1")) & "</td>")'
				'Response.Write("<td width='80' class='rText' align='right'>" & objRS4(("RCSAC" & recordNumber & "1" & recordNumber)) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & objRS4(("RCSAC" & "1" & recordNumber)) & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS4(("RCSAA" & recordNumber & "1")), 2) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS4(("RCSAA" & "1" & recordNumber)), 2) & "</td>")
				sesTotal = sesTotal + objRS4("RCSAA" & "1" & recordNumber)
				Response.Write("</tr>")
			printTaxValues4(recordNumber)
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
						'Response.Write("<td width='80' class='rText'>" & subrecordnumber & recordnumber & "</td>")
				tempsubrecordnumber = subrecordnumber
				subrecordNumber = 7
		End if
		Next
	'end if
	'Response.Write("<td width='80' class='rText'>RCSAA3+1" & subrecordnumber & recordnumber & "</td>")
	'Response.Write("<td width='80' class='rText'>&nbsp;</td>")
	End Function
'*******   Function added here to take the Table 4 records print them to the screen.
Function printTaxValues4(recordNumber)
	'Dim recordValue'
	Dim subrecordNumber, tempsubrecordnumber
	'For recordValue=2 to 6'
	For subrecordNumber=2 to 6
		'If objRS4(("RCAMT" & recordValue & "1")) > 0 Then'
		'Response.Write("the entry of printTaxValues 2-6 " & subrecordNumber )'
		If (objRS4(("RCSAA" & subrecordNumber & recordNumber )) > 0) or (objRS4(("RCAMT" & subrecordNumber & recordNumber))>0) Then
				Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText' align='right'>"  & objRS4(("RCTYP" & subrecordNumber & recordNumber)) & "</td>")
				If objRS4(("RCAMT" & subrecordNumber & recordNumber )) > 0 Then
					Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS4(("RCAMT" & subrecordNumber & recordNumber)), 2) & "</td>")
					sesTotal = sesTotal + objRS4("RCAMT" & subrecordNumber & recordNumber )
				Else
					Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				End If
				If objRS4("RCSAA" & subrecordNumber & recordNumber) > 0 Then
				Response.Write("<td width='80' class='rText' align='right'>" & objRS4(("RCSAR" & subrecordNumber & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & objRS4(("RCSAC" & subrecordNumber & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS4(("RCSAA" & subrecordNumber & recordNumber)), 2) & "</td>")
				Response.Write("</tr>")
				else
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				Response.Write("</tr>")

				End If
					sesTotal = sesTotal + objRS4("RCSAA" & subrecordNumber & recordNumber)
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

		If (objRS4(("RCSAA"& subrecordNumber & recordNumber )) = 0) and (objRS4(("RCAMT" & subrecordNumber & recordNumber))=0) Then
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
				If (objRS4(("RCSAA2" & recordNumber )) = 0) or (objRS4(("RCSAA3" & recordNumber )) = 0) Then
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

					If  (objRS4(("RCAMT2" & recordNumber )) = 0) and (objRS4(("RCSAA3" & recordNumber )) = 0) Then
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
						if (objRS4(("RCAMT2" & recordNumber + 1 )) = 0) or (objRS4(("RCSAA3" & recordNumber + 1 )) = 0) Then
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
'****  Ends the Function with RS4


'**    Begin the Function with RS6 here     **
Function printTaxRecord6()
	Dim recordNumber
	For recordNumber= 11 to 15
		If objRS6("RCBTDT" & recordNumber) = 0 Then
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
			if recordnumber < 10 then
			If objRS6("RCBTDT" & (recordNumber + 1)) = 0 Then
			'Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")'
			else
				'Response.Write("<td width='80' class='rText' align='right'>" & recordnumber & subrecordnumber & "</td>")
			end if
			end if
			'Response.Write("</tr>")
			recordNumber = 15
		Else
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='260' class='rText' align='left' rowspan='6'>")
				'Response.Write("<b>" & calcDateRS6(("RCBTDT" & recordNumber), (recordNumber)) & "</b><br>")'
				Response.Write("<b>" & calcDateRS6(("RCBTDT"), recordnumber) & "</b><br>")
				Response.Write("<b>Batch # " & objRS6(("RCBAT#" & recordNumber)) & "</b><br>")
				Response.Write("<b>Paid</b> by " & objRS6(("RCPDBY" & recordNumber)) & "<br>")
				Response.Write("<b>Validation #</b> " & objRS6(("RCVAL#" & recordNumber)))
				Response.Write("</td>")
				'Response.Write("<td width='60' class='rText' align='right'>" & objRS6(("RCTYP" & recordNumber & "1")) & "</td>")'
				Response.Write("<td width='60' class='rText' align='right'>" & objRS6(("RCTYP"  & "1" & recordNumber)) & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS6(("RCAMT" & recordNumber & "1")), 2) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS6(("RCAMT"  & "1" & recordNumber)), 2) & "</td>")
				sesTotal = objRS6("RCAMT" & "1" & recordNumber)
				'Response.Write("<td width='80' class='rText' align='right'>" & objRS6(("RCSAR" & recordNumber & "1")) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & objRS6(("RCSAR"  & "1" & recordNumber)) & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>" & objRS6(("RCSAC" & recordNumber & "1")) & "</td>")'
				'Response.Write("<td width='80' class='rText' align='right'>" & objRS6(("RCSAC" & recordNumber & "1" & recordNumber)) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & objRS6(("RCSAC" & "1" & recordNumber)) & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS6(("RCSAA" & recordNumber & "1")), 2) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS6(("RCSAA" & "1" & recordNumber)), 2) & "</td>")
				Session("recnumberend") = recordNumber
				sesTotal = sesTotal + objRS6("RCSAA" & "1" & recordNumber)
				Response.Write("</tr>")
			printTaxValues6(recordNumber)
		End If
	Next
End Function

'**    Ends the Function with RS6 here    **
'*******   Function added here to take the Table 6 records print them to the screen.
Function printTaxValues6(recordNumber)
	'Dim recordValue'
	Dim subrecordNumber, tempsubrecordnumber
	'For recordValue=2 to 6'
	For subrecordNumber=2 to 6
		'If objRS6(("RCAMT" & recordValue & "1")) > 0 Then'
		'Response.Write("the entry of printTaxValues 2-6 " & subrecordNumber )'
		If (objRS6(("RCSAA" & subrecordNumber & recordNumber )) > 0) or (objRS6(("RCAMT" & subrecordNumber & recordNumber))>0) Then
				Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText' align='right'>"  & objRS6(("RCTYP" & subrecordNumber & recordNumber)) & "</td>")
				If objRS6(("RCAMT" & subrecordNumber & recordNumber )) > 0 Then
					Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS6(("RCAMT" & subrecordNumber & recordNumber)), 2) & "</td>")
					sesTotal = sesTotal + objRS6("RCAMT" & subrecordNumber & recordNumber )
				Else
					Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				End If
				If objRS6("RCSAA" & subrecordNumber & recordNumber) > 0 Then
				Response.Write("<td width='80' class='rText' align='right'>" & objRS6(("RCSAR" & subrecordNumber & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & objRS6(("RCSAC" & subrecordNumber & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS6(("RCSAA" & subrecordNumber & recordNumber)), 2) & "</td>")
				Response.Write("</tr>")
				else
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				Response.Write("</tr>")

				End If
					sesTotal = sesTotal + objRS6("RCSAA" & subrecordNumber & recordNumber)
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

		If (objRS6(("RCSAA"& subrecordNumber & recordNumber )) = 0) and (objRS6(("RCAMT" & subrecordNumber & recordNumber))=0) Then
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
				If (objRS6(("RCSAA2" & recordNumber )) = 0) or (objRS6(("RCSAA3" & recordNumber )) = 0) Then
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

					If  (objRS6(("RCAMT2" & recordNumber )) = 0) and (objRS6(("RCSAA3" & recordNumber )) = 0) Then
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
						if (objRS6(("RCAMT2" & recordNumber + 1 )) = 0) or (objRS6(("RCSAA3" & recordNumber + 1 )) = 0) Then
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
'****  Ends the Function with RS6

'**    Begin the Function with RS7 here     **
Function printTaxRecord7()
	Dim recordNumber
	For recordNumber= 15 to 20
		If objRS7("RCBTDT" & recordNumber) = 0 Then
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
			if recordnumber < 10 then
			If objRS7("RCBTDT" & (recordNumber + 1)) = 0 Then
			'Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")'
			else
				'Response.Write("<td width='80' class='rText' align='right'>" & recordnumber & subrecordnumber & "</td>")
			end if
			end if
			'Response.Write("</tr>")
			recordNumber = 20
		Else
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='260' class='rText' align='left' rowspan='6'>")
				'Response.Write("<b>" & calcDateRS7(("RCBTDT" & recordNumber), (recordNumber)) & "</b><br>")'
				Response.Write("<b>" & calcDateRS7(("RCBTDT"), recordnumber) & "</b><br>")
				Response.Write("<b>Batch # " & objRS7(("RCBAT#" & recordNumber)) & "</b><br>")
				Response.Write("<b>Paid</b> by " & objRS7(("RCPDBY" & recordNumber)) & "<br>")
				Response.Write("<b>Validation #</b> " & objRS7(("RCVAL#" & recordNumber)))
				Response.Write("</td>")
				'Response.Write("<td width='60' class='rText' align='right'>" & objRS7(("RCTYP" & recordNumber & "1")) & "</td>")'
				Response.Write("<td width='60' class='rText' align='right'>" & objRS7(("RCTYP"  & "1" & recordNumber)) & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS7(("RCAMT" & recordNumber & "1")), 2) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS7(("RCAMT"  & "1" & recordNumber)), 2) & "</td>")
				sesTotal = objRS7("RCAMT" & "1" & recordNumber)
				'Response.Write("<td width='80' class='rText' align='right'>" & objRS7(("RCSAR" & recordNumber & "1")) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & objRS7(("RCSAR"  & "1" & recordNumber)) & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>" & objRS7(("RCSAC" & recordNumber & "1")) & "</td>")'
				'Response.Write("<td width='80' class='rText' align='right'>" & objRS7(("RCSAC" & recordNumber & "1" & recordNumber)) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & objRS7(("RCSAC" & "1" & recordNumber)) & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS7(("RCSAA" & recordNumber & "1")), 2) & "</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS7(("RCSAA" & "1" & recordNumber)), 2) & "</td>")
				Session("recnumberend") = recordNumber
				sesTotal = sesTotal + objRS7("RCSAA" & "1" & recordNumber)
				Response.Write("</tr>")
			printTaxValues7(recordNumber)
		End If
	Next
End Function

'**    Ends the Function with RS7 here    **

'*******   Function added here to take the Table 6 records print them to the screen.
Function printTaxValues7(recordNumber)
	'Dim recordValue'
	Dim subrecordNumber, tempsubrecordnumber
	'For recordValue=2 to 6'
	For subrecordNumber=2 to 6
		'If objRS7(("RCAMT" & recordValue & "1")) > 0 Then'
		'Response.Write("the entry of printTaxValues 2-6 " & subrecordNumber )'
		If (objRS7(("RCSAA" & subrecordNumber & recordNumber )) > 0) or (objRS7(("RCAMT" & subrecordNumber & recordNumber))>0) Then
				Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText' align='right'>"  & objRS7(("RCTYP" & subrecordNumber & recordNumber)) & "</td>")
				If objRS7(("RCAMT" & subrecordNumber & recordNumber )) > 0 Then
					Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS7(("RCAMT" & subrecordNumber & recordNumber)), 2) & "</td>")
					sesTotal = sesTotal + objRS6("RCAMT" & subrecordNumber & recordNumber )
				Else
					Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				End If
				If objRS7("RCSAA" & subrecordNumber & recordNumber) > 0 Then
				Response.Write("<td width='80' class='rText' align='right'>" & objRS7(("RCSAR" & subrecordNumber & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & objRS7(("RCSAC" & subrecordNumber & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS7(("RCSAA" & subrecordNumber & recordNumber)), 2) & "</td>")
				Response.Write("</tr>")
				else
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				Response.Write("</tr>")

				End If
					sesTotal = sesTotal + objRS7("RCSAA" & subrecordNumber & recordNumber)
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

		If (objRS7(("RCSAA"& subrecordNumber & recordNumber )) = 0) and (objRS7(("RCAMT" & subrecordNumber & recordNumber))=0) Then
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
				If (objRS7(("RCSAA2" & recordNumber )) = 0) or (objRS7(("RCSAA3" & recordNumber )) = 0) Then
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

					If  (objRS7(("RCAMT2" & recordNumber )) = 0) and (objRS7(("RCSAA3" & recordNumber )) = 0) Then
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
						if (objRS7(("RCAMT2" & recordNumber + 1 )) = 0) or (objRS7(("RCSAA3" & recordNumber + 1 )) = 0) Then
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
'****  Ends the Function with RS7




Function calcZero(strData, intPlaces)
	If objRS(strData) <> "" Then
		strData = (FormatNumber(objRS(strData), intPlaces))
	Else
		strData = "0.00"
	End If
	calcZero = strData
end Function

Function calcZeroPY(strData, intPlaces)
	If objRSPY(strData) <> "" Then
		strData = (FormatNumber(objRSPY(strData), intPlaces))
	Else
		strData = "0.00"
	End If
	calcZeroPY = strData
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

Function calcDateRS6(strData, intrecordnumber)
	Dim strYear, strMonth, strDay, intLength
	strData = objRS6("RCBTDT" & intrecordnumber)
	intLength = Len(strData)
	strYear = Right(strData, 4)
	If intLength = 7 Then
		strMonth = Left(strData, 1)
		strDay = Mid(strData, 2, 2)
	Else
		strMonth = Left(strData, 2)
		strDay = Mid(strData, 3, 2)
	End If
	calcDateRS6 = strMonth + "/" + strDay + "/" + strYear
end Function

Function calcDateRS7(strData, intrecordnumber)
	Dim strYear, strMonth, strDay, intLength
	strData = objRS7("RCBTDT" & intrecordnumber)
	intLength = Len(strData)
	strYear = Right(strData, 4)
	If intLength = 7 Then
		strMonth = Left(strData, 1)
		strDay = Mid(strData, 2, 2)
	Else
		strMonth = Left(strData, 2)
		strDay = Mid(strData, 3, 2)
	End If
	calcDateRS7 = strMonth + "/" + strDay + "/" + strYear
end Function


Function calcDateRS10(strData, intrecordnumber)
	Dim strYear, strMonth, strDay, intLength
	strData = objRS10("SRAUDT" & intrecordnumber)
	intLength = Len(strData)
	strYear = Right(strData, 4)
	If intLength = 7 Then
		strMonth = Left(strData, 1)
		strDay = Mid(strData, 2, 2)
	Else
		strMonth = Left(strData, 2)
		strDay = Mid(strData, 3, 2)
	End If
	calcDateRS10 = strMonth + "/" + strDay + "/" + strYear
end Function

Function calcDate6Digit(strData, intrecordnumber)
	Dim strYear, strMonth, strDay, intLength
	strData = objRS10("SRSLDT" & intrecordnumber)
	intLength = Len(strData)
	strYear = Right(strData, 4)
	If intLength = 5 Then
		strMonth = Left(strData, 1)
	Else
		strMonth = Left(strData, 2)
	End If
	calcDate6Digit = strMonth + "/" + strYear
end Function

Function calcZip(strData)
	If objRS(strData) = "00000" Then
		strData = ""
	End If
	calcZip = strData
end Function

Dim objCommand, objRS, strQueryString, strYEAR, strPID, strTID, objRS3, objRS5, objRSV, objRSV2, objRS10, objRScount, intnumberQ
Dim intYYR, intYEAR

'intYEAR = 2005
Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("YEAR")
strQueryString = request.QueryString("pid")
'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strQueryString & "' AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR="& ((intYYR)-1) &";"
'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strQueryString & "' AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR="& ((intYYR)-1) &"  AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG='T' ;"
strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strQueryString & "' AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR="& ((intYYR)) &"  AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG='V' ;"
'Response.Write("The value of the Query String : " & strQueryString )

'Fill in the command properties
'Response.Write( "the strConnect is " & strConnect   )
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
objCommand.CommandType = 1

Set objRSPY = objCommand.Execute

Set objCommand = Nothing

'create a record set to handle when the user takes a 'Value Year' and then there needs to show a History
'New RecordSet to handle History
Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("YEAR")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strQueryString & "' AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR="& intYYR &";"
'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strQueryString & "' AND WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR="& intYYR &"  AND WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG='V' ;"
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


'TEST AREA * * * * * * * * * * * * * * * * * * * * * * ** * * * * * * *
'set conn=Server.CreateObject("ADODB.Connection")
'conn.Provider="Microsoft.Jet.OLEDB.4.0"
'conn.Open(Server.Mappath("WebTab67.mdb"))
'set rs = Server.CreateObject("ADODB.recordset")
'strPID = request.QueryString("pid")
'strTID = request.QueryString("tid")
'strRID = request.QueryString("rid")
'strQueryString = request.QueryString("pid")
''response.write(" value for strRID : " & strRID )
''response.write("  value for strQueryString : " & strQueryString )
'if strRID = "" then
''strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' AND [Table 9 - Value Info].YEAR = " & intYR & " AND [Table 9 - Value Info].RecNum =3 ;"
'strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' AND [Table 9 - Value Info].YEAR = " & intYR & " ORDER BY [Table 9 - Value Info].RecNum;"
'else
'strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' AND [Table 9 - Value Info].YEAR = " & intYR & " AND [Table 9 - Value Info].RecNum = " & strRID & " ;"
''strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' AND [Table 9 - Value Info].YEAR = " & intYR & ";"
''strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' AND [Table 9 - Value Info].YEAR = " & intYR & " AND [Table 9 - Value Info].RecNum = " & Session("intREC") & " ;"
'end if
'
'sql="SELECT * FROM [Table 9 - Value Info] " & strQueryString
''response.write("the strQueryString is : " & strQueryString )
''response.write("value for sql in rs : " & sql )
'rs.Open sql,conn,3,1
''rs.MoveFirst
''rs.MoveNext
''rs.MoveNext
'*&*&*&*&*(&*&**((((**&&********************************8

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strRID = request.QueryString("rid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
if strRID = "" then
'strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' AND [Table 9 - Value Info].YEAR = " & intYYR & " AND [Table 9 - Value Info].RecNum =3 ;"
strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' AND [Table 9 - Value Info].YEAR = " & intYYR & " ORDER BY [Table 9 - Value Info].RecNum;"
else
strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' AND [Table 9 - Value Info].YEAR = " & intYYR & " AND [Table 9 - Value Info].RecNum = " & strRID & " ;"
'strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' AND [Table 9 - Value Info].YEAR = " & intYYR & ";"
'strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' AND [Table 9 - Value Info].YEAR = " & intYYR & " AND [Table 9 - Value Info].RecNum = " & Session("intREC") & " ;"
end if
'Fill in the command properties
objCommand.ActiveConnection = strConnect
'set conn=Server.CreateObject("ADODB.Connection")
'conn.Provider="Microsoft.Jet.OLEDB.4.0"
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1
'set rs = Server.CreateObject("ADODB.recordset")
'rs.CursorType = adOpenStatic
'rs.LockType=adLockReadOnly
Set rs = objCommand.Execute
'rs.Open objCommand.CommandText,conn,3,1
'rs.Open sql,conn,3,1
'rs.CursorType = adOpenStatic

Set objCommand = Nothing

'   END TEST   AREA     * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 3 - Part1 of RCPT Sets(1-5)].TXPRCL = '" & strQueryString & "'"

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 3 - Part1 of RCPT Sets(1-5)] " & strQueryString
objCommand.CommandType = 1

Set objRS3 = objCommand.Execute
Set objCommand = Nothing

' To create RS4 from Table 4 Tax Receipts

Set objCommand = Server.CreateObject("ADODB.Command")
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 4 - Part2 of RCPT Sets(6-10)].TXPRCL = '" & strQueryString & "'"

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 4 - Part2 of RCPT Sets(6-10)] " & strQueryString
objCommand.CommandType = 1

Set objRS4 = objCommand.Execute
Set objCommand = Nothing

'** ends table 4 create

Set objCommand = Nothing

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 5 - Unpaid Taxes/Truth in Taxation Info].TXPRCL = '" & strQueryString & "'"

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 5 - Unpaid Taxes/Truth in Taxation Info] " & strQueryString
objCommand.CommandType = 1

Set objRS5 = objCommand.Execute
Set objCommand = Nothing
'CREATE THE OBJRSV FROM TABLE 9 **********************

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' AND [Table 9 - Value Info].YEAR =  "& intYYR &""

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRSV = objCommand.Execute
intYEAR = objRSV("Year")

Set objCommand = Nothing
'CREATE THE OBJRSSP *****************
Set objCommand = Server.CreateObject("ADODB.Command")
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 2 - Special/Ditch Info].TXPRCL = '" & strQueryString & "' AND [Table 2 - Special/Ditch Info].TXYEAR =" & intYYR & " "
'strQueryString = "WHERE [Table 2 - Special/Ditch Info Query].TXPRCL = '" & strQueryString & "' AND [Table 2 - Special/Ditch Info Query].TXYEAR =intYEAR; "
'strQueryString = "WHERE [Table 2 - Special/Ditch Info Query].TXPRCL = '" & strQueryString & "' AND [Table 2 - Special/Ditch Info Query].TXYEAR =[intYEAR];"
'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 2 - Special/Ditch Info] " & strQueryString
'objCommand.CommandText = "SELECT * FROM [Table 2 - Special/Ditch Info Query] " & strQueryString
objCommand.CommandType = 1
'Response.Write(" the value of strQueryString : " & strQueryString )
Set objRSSP = objCommand.Execute
Set objCommand = Nothing

'CREATE THE OBJRS10   *****************
Set objCommand = Server.CreateObject("ADODB.Command")
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 10 - Sales Info].PARCEL = '" & strQueryString & "' "
'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 10 - Sales Info] " & strQueryString
objCommand.CommandType = 1
'Response.Write(" the value of strQueryString : " & strQueryString )
Set objRS10 = objCommand.Execute
Set objCommand = Nothing
'CREATE THE OBJRSV2 FROM TABLE 9 **********************

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strRID = request.QueryString("rid")
intYYR = request.QueryString("yr")
'response.write("the value for intYYR : " & intYYR )
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' ORDER BY [Table 9 - Value Info].Parcel, [Table 9 - Value Info].Year DESC, [Table 9 - Value Info].RecNum DESC;"

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRSV2 = objCommand.Execute
Set objCommand = Nothing
'Response.Write(" value for intRC :" & intRC & "Value for SessionTOTREC : " & totrec )
'Response.Write("Value for session intREC : " & Session("intREC"))
'Response.Write("Value for strRID : " & strRID )

Session("intREC") = rs("RECNUM")
'if strRID <> "" Then
'if strRID = Session("intREC") Then
'rs.MoveNext
''Session("intREC") = rs("RECNUM")
''intRC = rs("RECNUM")
'intRC = Session("intREC")
'else
'rs.MoveNext
''Session("intREC") = rs("RECNUM")
'
'intRC = Session("intREC")
''intRC = rs("RECNUM")
'end if
'end if
Session("TOTREC") = objRSV2("RecNum")
%>

<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="650" align="right" colspan="2">Parcel Number: <b><%= objRSV("PARCEL") %></b></td>
	</tr>
	<tr valign="top">

<%
'code pasted in here for a test to use the 'Total' for record  identification to user.
		if objRSV2("RecNum") = 1 then
			'Response.Write("<td width='650' align='right' colspan='2'>Value Year: <b>" & objRSV2("YEAR") & " Rec# Total </b></td>")
			'Response.Write("<td width='650' align='right' colspan='2'>Value Year: <b> " & objRS("TXYEAR") & " Rec# " & rs("RecNum") & " of " & objRSV2("RecNum") & "</b></td>")
			Response.Write("<td width='650' align='right' colspan='2'>Payable Year: <b> " & intYYR & " Rec# " & rs("RecNum") & " of " & objRSV2("RecNum") & "</b></td>")
		else
			if strRID < objRSV2("RecNum") Then
				If strRID = "" then
				'rs.MoveNext
				else
					Response.Write("<td width='650' align='right' colspan='2'>Payable Year: <b> " & intYYR & " Rec# " & rs("RecNum") & " of " & objRSV2("RecNum") & "</b></td>")
					'Response.Write("<td width='650' align='right' colspan='2'>Value Year: <b>" & intYYR & " Rec# Total </b></td>")
				end if
				if strRID = rid then
				else
				end if
			else
				if (rs("RecNum")) = (objRSV2("RecNum")) Then
					Response.Write("<td width='650' align='right' colspan='2'>Payable Year: <b> " & intYYR & " Rec# " & rs("RecNum") & " of " & objRSV2("RecNum") & "</b></td>")
					if strRID = 0 then
					Response.Write("<td width='650' align='right' colspan='2'>Payable Year: <b>" & intYYR & " Rec# Total </b></td>")
					end if
				else
					if strRID = 0 then
					'Response.Write("<td width='650' align='right' colspan='2'>Value Year: <b> " & intYYR & " Rec# " & rs("RecNum") & " of " & objRSV2("RecNum") & "</b></td>")
					Response.Write("<td width='650' align='right' colspan='2'>Payable Year: <b>" & intYYR & " Rec# Total </b></td>")
					else
					Response.Write("<td width='650' align='right' colspan='2'>Payable Year: <b> " & intYYR & " Rec# " & rs("RecNum") & " of " & objRSV2("RecNum") & "</b></td>")
					end if
				end if
			end if
		end if


%>
	<tr valign="top">
		<td width="650" align="right" colspan="2"> <b>
		<%
' code to go to the recnum=1 not to recnum=0 . Recnum = 0 will be handled differently.
'Recnum = 0 is the Total value of all the recnums when there is more than one record in Table 9.
	'if rs("RecNum") = 0
		if objRSV2("RecNum") = 1 then
		'Response.write(" the value for the strRID : " & strRID )
		'Response.Write("<a class='pLink' href='Parcel_Rock.asp?pid=" & strPID & "&tid=1&rid="& strRID + 1 &"'>Next Record</a>")
		else
		'Response.write(" the value for the strRID : " & strRID )
		'Response.write(" the value for the Session(TOTREC) :  " & Session("TOTREC") )
		'Response.Write(" if objRSV2(recnum) = 1 the else : ")
		'Response.write(" the value for the objRSV2(RecNum) :  " & objRSV2("RecNum") )
			if strRID < objRSV2("RecNum") Then
			'Response.Write(" if strRID < objRSV2(recnum)  : ")
				If strRID = "" then
				strRID = 0
				Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID + 1 &"'>Next Record</a>")
				else
				Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID + 1 &"'>Next Record</a>")
					if strRID >= 1 Then
					Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID - 1 &"'>| Previous Record</a>")
					Response.Write(" the value for the objRSV2(RecNum) :  " & objRSV2("RecNum") )
					else
						'if strRID = 0 then
						'Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid=" & strRID & "'>Total Value</a>")
						'Response.Write(" the  value for the rid : " & rid )
						'end if
					end if
				end if
				'Response.write(" the value for the strRID : " & strRID )
				'Response.Write(" the  value for the rid : " & rid )
				'Response.Write(" the value for the totrec : " & totrec )

				if strRID = rid then
				else
					'objRSV.MoveNext

					rs.MoveNext
					'strRID = strRID + 1
				end if
			else
				'if strRID = objRSV2("RecNum") Then
				if (rs("RecNum")) = (objRSV2("RecNum")) Then
				'Response.Write(" the two are equal  ")
				Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid=0&yr="& objRS("TXYEAR") &"'>Total </a>")
				Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID - 1 &"&yr="& objRS("TXYEAR") &"'>| Previous Record</a>")
				else
				'Response.Write(" the  value for the rid : " & rid )
				'Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID + 1 &"&yr="& objRS("TXYEAR") &"'>Next Record</a>")

					if strRID >= 1 Then
					Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID + 1 &"&yr="& objRS("TXYEAR") &"'>Next Record</a>")
					Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID - 1 &"&yr="& objRS("TXYEAR") &"'>| Previous Record</a>")
					end if
					if strRID = 0 then
					Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid=" & objRSV2("RecNum") & "&yr="& objRS("TXYEAR") &"'>| Previous Record</a>")
					'Response.Write(" the  value for the rid : " & rid )
					end if
				'Response.write(" the else of the strRID < objRSV2('RecNum')  statement ::")
				'Response.write(" the value for the strRID : " & strRID )
				'Response.write(" the value for the objRSV2(RecNum) :  " & Session("TOTREC") )
				' this ends the forward read of the rs recordset and showing the 'NEXT'
				end if
			end if

		end if


		%></b></td>
	</tr>
	<tr>
		<td width="10"></td>
		<td width="650" align="left">
		<%
			Select Case strTID
			Case 0
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=0&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='hLink'>General Information</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Value Information</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=2&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Special Amounts</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=3&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Ditch</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=4&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Sales</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=5&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>History</a>")


			Case 1
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=0&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='hLink'>Value Information</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=2&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Special Amounts</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=3&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Ditch</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=4&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Sales</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=5&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>History</a>")

			Case 2
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=0&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Value Information</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=2&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='hLink'>Special Amounts</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=3&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Ditch</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=4&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Sales</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=5&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>History</a>")

			Case 3
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=0&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Value Information</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=2&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Special Amounts</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=3&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='hLink'>Ditch</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=4&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Sales</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=5&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>History</a>")

			Case 4
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=0&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Value Information</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=2&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Special Amounts</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=3&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Ditch</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=4&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='hLink'>Sales</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=5&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>History</a>")

			Case 5
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=0&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>General Information</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Value Information</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=2&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Special Amounts</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=3&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Ditch</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=4&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='uLink'>Sales</a>|")
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=5&rid="& strRID &"&yr="& objRS("TXYEAR") &"' class='hLink'>History</a>")

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
		<!-- #include file="GeneralTax_Rock.asp" -->
		<%
			Case 1
		%>
		<!-- #include file="ValueInformation.asp" -->
		<%
			Case 2
		%>
		<!-- #include file="SpecialAssessment.asp" -->
		<%
			Case 3
		%>
		<!-- #include file="DitchAssessment.asp" -->
		<%
			Case 4
		%>
		<!-- #include file="Sales.asp" -->
		<%
			Case 5
		%>
		<!-- #include file="History_Rock.asp" -->
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

	Response.Write("<a href='/tax/searchinputreturn_Rock.asp' class='tlink'>Another Search    |</a>&nbsp;&nbsp;&nbsp;&nbsp")
	Response.Write("<a href='/tax/ParcelListReturn_Rock.asp' class='tlink'>Back to ParcelList    |</a>&nbsp;&nbsp;&nbsp;&nbsp")
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




'objRS.Close
'Set objRS = Nothing
'objRSV.Close
'Set objRSV = Nothing
'objRS3.Close
'Set objRS3 = Nothing
'objRS5.Close
'Set objRS5 = Nothing

'objRScount.Close
'Set objRScount = Nothing

%>


</body>
</html>
