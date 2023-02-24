<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Parcel Search Results</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
	Dim cid, cName, sesTotal, intYR, intRC
	cid = request.QueryString("cid")
	cName = Session("County Name")
	sesTotal = Session("amtTotal")
	intYR = Session("intYEAR")
	totrec = Session("TOTREC")

	intParcelNo = request.QueryString("varintParcelNo")
	strAddress = request.QueryString("varstrAddress")
	strName = request.QueryString("varstrName")
	intSect = request.QueryString("varintSect")
	intTwp = request.QueryString("varintTwp")
	intRange = request.QueryString("varintRange")


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
				End If
			recordNumber = 5
		Else
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='260' class='rText' align='left' rowspan='6'>")
				Response.Write("<b>" & calcDate(("RCBTDT"), recordnumber) & "</b><br>")
				Response.Write("<b>Batch # " & objRS3(("RCBAT#" & recordNumber)) & "</b><br>")
				Response.Write("<b>Paid</b> by " & objRS3(("RCPDBY" & recordNumber)) & "<br>")
				Response.Write("<b>Validation #</b> " & objRS3(("RCVAL#" & recordNumber)))
				Response.Write("</td>")
				Response.Write("<td width='60' class='rText' align='right'>" & objRS3(("RCTYP"  & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCAMT"  & "1" & recordNumber)), 2) & "</td>")
				sesTotal = objRS3("RCAMT" & "1" & recordNumber)
				Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAR"  & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAC" & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCSAA" & "1" & recordNumber)), 2) & "</td>")
				sesTotal = sesTotal + objRS3("RCSAA" & "1" & recordNumber)
				Response.Write("</tr>")
			printTaxValues(recordNumber)
		End If
	Next
End Function


Function printTaxRecord4()
	Dim recordNumber
	For recordNumber=1 to 5
		If objRS4("RCBTDT" & recordNumber) = 0 Then
			Response.Write("<tr valign='top'>")
				If recordNumber = 1 Then
					Response.Write("<td class='rText' align='left' rowspan='7'>No Tax Receipt Information</td>")
				End If
			recordNumber = 5
		Else
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='260' class='rText' align='left' rowspan='6'>")
				Response.Write("<b>" & calcDate(("RCBTDT"), recordnumber) & "</b><br>")
				Response.Write("<b>Batch # " & objRS4(("RCBAT#" & recordNumber)) & "</b><br>")
				Response.Write("<b>Paid</b> by " & objRS4(("RCPDBY" & recordNumber)) & "<br>")
				Response.Write("<b>Validation #</b> " & objRS4(("RCVAL#" & recordNumber)))
				Response.Write("</td>")
				Response.Write("<td width='60' class='rText' align='right'>" & objRS4(("RCTYP"  & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS4(("RCAMT"  & "1" & recordNumber)), 2) & "</td>")
				sesTotal = objRS4("RCAMT" & "1" & recordNumber)
				Response.Write("<td width='80' class='rText' align='right'>" & objRS4(("RCSAR"  & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & objRS4(("RCSAC" & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS4(("RCSAA" & "1" & recordNumber)), 2) & "</td>")
				sesTotal = sesTotal + objRS4("RCSAA" & "1" & recordNumber)
				Response.Write("</tr>")
			printTaxValues4(recordNumber)
		End If
	Next
End Function


Function printTaxValues(recordNumber)
	Dim subrecordNumber, tempsubrecordnumber
	For subrecordNumber=2 to 6
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
				Response.Write("</tr>")
		End If

		If (objRS3(("RCSAA"& subrecordNumber & recordNumber )) = 0) and (objRS3(("RCAMT" & subrecordNumber & recordNumber))=0) Then
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("</tr>")

			Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("</tr>")

				Response.Write("<tr valign='top'>")
					Response.Write("<td width='60' class='rText'>&nbsp;</td>")
					Response.Write("<td width='80' class='rText'>&nbsp;</td>")
					Response.Write("<td width='80' class='rText'>&nbsp;</td>")
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

					If  (objRS3(("RCAMT2" & recordNumber )) = 0) and (objRS3(("RCSAA3" & recordNumber )) = 0) Then
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						if subrecordnumber > 2 then
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						end if
					else
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						if (objRS3(("RCAMT2" & recordNumber + 1 )) = 0) or (objRS3(("RCSAA3" & recordNumber + 1 )) = 0) Then
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						else
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						end if
					end if
						Response.Write("</tr>")
				tempsubrecordnumber = subrecordnumber
				subrecordNumber = 7
		End if
		Next
	End Function
'*******   Function added here to take the Table 4 records print them to the screen.
Function printTaxValues4(recordNumber)
	Dim subrecordNumber, tempsubrecordnumber
	For subrecordNumber=2 to 6
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
				Response.Write("</tr>")
		End If

		If (objRS4(("RCSAA"& subrecordNumber & recordNumber )) = 0) and (objRS4(("RCAMT" & subrecordNumber & recordNumber))=0) Then
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("</tr>")

			Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("</tr>")

				Response.Write("<tr valign='top'>")
					Response.Write("<td width='60' class='rText'>&nbsp;</td>")
					Response.Write("<td width='80' class='rText'>&nbsp;</td>")
					Response.Write("<td width='80' class='rText'>&nbsp;</td>")
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
					If  (objRS4(("RCAMT2" & recordNumber )) = 0) and (objRS4(("RCSAA3" & recordNumber )) = 0) Then
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						if subrecordnumber > 2 then
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						end if
					else
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						if (objRS4(("RCAMT2" & recordNumber + 1 )) = 0) or (objRS4(("RCSAA3" & recordNumber + 1 )) = 0) Then
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						else
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						end if
					end if
						Response.Write("</tr>")
				tempsubrecordnumber = subrecordnumber
				subrecordNumber = 7
		End if
		Next
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
				End If
			recordNumber = 15
		Else
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='260' class='rText' align='left' rowspan='6'>")
				Response.Write("<b>" & calcDateRS6(("RCBTDT"), recordnumber) & "</b><br>")
				Response.Write("<b>Batch # " & objRS6(("RCBAT#" & recordNumber)) & "</b><br>")
				Response.Write("<b>Paid</b> by " & objRS6(("RCPDBY" & recordNumber)) & "<br>")
				Response.Write("<b>Validation #</b> " & objRS6(("RCVAL#" & recordNumber)))
				Response.Write("</td>")
				Response.Write("<td width='60' class='rText' align='right'>" & objRS6(("RCTYP"  & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS6(("RCAMT"  & "1" & recordNumber)), 2) & "</td>")
				sesTotal = objRS6("RCAMT" & "1" & recordNumber)
				Response.Write("<td width='80' class='rText' align='right'>" & objRS6(("RCSAR"  & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & objRS6(("RCSAC" & "1" & recordNumber)) & "</td>")
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
	Dim subrecordNumber, tempsubrecordnumber
	For subrecordNumber=2 to 6
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
				Response.Write("</tr>")
		End If

		If (objRS6(("RCSAA"& subrecordNumber & recordNumber )) = 0) and (objRS6(("RCAMT" & subrecordNumber & recordNumber))=0) Then
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("</tr>")

			Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("</tr>")

				Response.Write("<tr valign='top'>")
					Response.Write("<td width='60' class='rText'>&nbsp;</td>")
					Response.Write("<td width='80' class='rText'>&nbsp;</td>")
					Response.Write("<td width='80' class='rText'>&nbsp;</td>")
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

					If  (objRS6(("RCAMT2" & recordNumber )) = 0) and (objRS6(("RCSAA3" & recordNumber )) = 0) Then
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						if subrecordnumber > 2 then
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						end if
					else
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						if (objRS6(("RCAMT2" & recordNumber + 1 )) = 0) or (objRS6(("RCSAA3" & recordNumber + 1 )) = 0) Then
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						else
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						end if
					end if
						Response.Write("</tr>")
				tempsubrecordnumber = subrecordnumber
				subrecordNumber = 7
		End if
		Next
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
				End If
			recordNumber = 20
		Else
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='260' class='rText' align='left' rowspan='6'>")
				Response.Write("<b>" & calcDateRS7(("RCBTDT"), recordnumber) & "</b><br>")
				Response.Write("<b>Batch # " & objRS7(("RCBAT#" & recordNumber)) & "</b><br>")
				Response.Write("<b>Paid</b> by " & objRS7(("RCPDBY" & recordNumber)) & "<br>")
				Response.Write("<b>Validation #</b> " & objRS7(("RCVAL#" & recordNumber)))
				Response.Write("</td>")
				Response.Write("<td width='60' class='rText' align='right'>" & objRS7(("RCTYP"  & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS7(("RCAMT"  & "1" & recordNumber)), 2) & "</td>")
				sesTotal = objRS7("RCAMT" & "1" & recordNumber)
				Response.Write("<td width='80' class='rText' align='right'>" & objRS7(("RCSAR"  & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & objRS7(("RCSAC" & "1" & recordNumber)) & "</td>")
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
	Dim subrecordNumber, tempsubrecordnumber
	For subrecordNumber=2 to 6
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
				Response.Write("</tr>")
		End If

		If (objRS7(("RCSAA"& subrecordNumber & recordNumber )) = 0) and (objRS7(("RCAMT" & subrecordNumber & recordNumber))=0) Then
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("</tr>")

			Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			Response.Write("</tr>")

				Response.Write("<tr valign='top'>")
					Response.Write("<td width='60' class='rText'>&nbsp;</td>")
					Response.Write("<td width='80' class='rText'>&nbsp;</td>")
					Response.Write("<td width='80' class='rText'>&nbsp;</td>")
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

					If  (objRS7(("RCAMT2" & recordNumber )) = 0) and (objRS7(("RCSAA3" & recordNumber )) = 0) Then
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						if subrecordnumber > 2 then
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						end if
					else
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						if (objRS7(("RCAMT2" & recordNumber + 1 )) = 0) or (objRS7(("RCSAA3" & recordNumber + 1 )) = 0) Then
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						else
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						end if
					end if
						Response.Write("</tr>")
				tempsubrecordnumber = subrecordnumber
				subrecordNumber = 7
		End if
		Next
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

Function calcZeroRSPYGEN(strData, intPlaces)
	If objRSPYGEN(strData) <> "" Then
		strData = (FormatNumber(objRSPYGEN(strData), intPlaces))
	Else
		strData = "0.00"
	End If
	calcZeroRSPYGEN = strData
end Function

Function calcZeroRS(strData, intPlaces)
	If objRSCT10(strData) <> 0 Then
		strData = (FormatNumber(objRSCT10(strData), intPlaces))
	Else
		strData = "0.00"
	End If
	calcZeroRS = strData
end Function

Function calcZeroRSC(strData, intPlaces)
	If objRSC10(strData) <> 0 Then
		strData = (FormatNumber(objRSC10(strData), intPlaces))
	Else
		strData = "0.00"
	End If
	calcZeroRSC = strData
end Function

Function calcZeroPY(strData, intPlaces)
	If objRSPY(strData) <> "" Then
		strData = (FormatNumber(objRSPY(strData), intPlaces))
	Else
		strData = "0.00"
	End If
	calcZeroPY = strData
end Function

Function calcZeroRSSP(strData, intPlaces)
	If objRSSP(strData) <> "" Then
		strData = (FormatNumber(objRSSP(strData), intPlaces))
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


'**  Function to format the entry of the parcel number search from the user
'**  this will allow the user to enter numbers only in the text field of the form.
'**  this is in ND web tax
Function fmtparcelnum(strData)
	Dim strtrimdata
	strData = Request.Form("ParcelNo")
	strtrimdata = Trim(strData)
	intLength = Len(strtrimdata)
If cid = 47 or cid = 30 or cid = 31 or cid = 34 or cid = 37 or cid = 23 or cid = 41 or cid = "02" Then '2-7
	If intLength = 1 Then
		fmtparcelnum = strtrimdata
	elseif intlength = 2 Then
		fmtparcelnum = strtrimdata
	elseif intlength = 3 Then
		strleftChars = Left(strtrimdata, 2)
		strrightChars = Right(strtrimdata, 1)
		fmtparcelnum = strleftChars + "-" + strrightChars
	elseif intlength = 4 Then
		strleftChars = Left(strtrimdata, 2)
		strrightChars = Right(strtrimdata, 2)
		fmtparcelnum = strleftChars + "-" + strrightChars
	elseif intlength = 5 Then
		strleftChars = Left(strtrimdata, 2)
		strrightChars = Right(strtrimdata, 3)
		fmtparcelnum = strleftChars + "-" + strrightChars
	elseif intlength = 6 Then
		strleftChars = Left(strtrimdata, 2)
		strrightChars = Right(strtrimdata, 4)
		fmtparcelnum = strleftChars + "-" + strrightChars
	elseif intLength = 7 Then
		strleftChars = Left(strtrimdata, 2)
		strrightChars = Right(strtrimdata, 5)
		fmtparcelnum = strleftChars + "-" + strrightChars
	elseif intLength = 8 Then
		strleftChars = Left(strtrimdata, 2)
		strrightChars = Right(strtrimdata, 6)
		fmtparcelnum = strleftChars + "-" + strrightChars
	elseif intlength = 9 Then
		strleftChars = Left(strtrimdata, 2)
		strrightChars = Right(strtrimdata, 7)
		fmtparcelnum = strleftChars + "-" + strrightChars
	End If
End If
If cid = 27  Then ' 2-2-5
		If intLength = 1 Then
			fmtparcelnum = strtrimdata
		elseif intlength = 2 Then
			fmtparcelnum = strtrimdata
		elseif intlength = 3 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Right(strtrimdata, 1)
			fmtparcelnum = strleftChars + "-" + strmidChars
		elseif intlength = 4 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Right(strtrimdata, 2)
			fmtparcelnum = strleftChars + "-" + strmidChars
		elseif intlength = 5 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 2)
			strrightChars = Right(strtrimdata, 1)
			fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
		'Response.write(" the value of the fmtparcelnum in len 5 : " & fmtparcelnum)'
		elseif intlength = 6 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 2)
			strrightChars = Right(strtrimdata, 2)
			fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
		'Response.write(" the value of the fmtparcelnum in len 6 : " & fmtparcelnum)'
		elseif intLength = 7 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 2)
			strrightChars = Right(strtrimdata, 3)
			fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
		elseif intLength = 8 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 2)
			strrightChars = Right(strtrimdata, 4)
			fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
		elseif intlength = 9 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 2)
			strrightChars = Right(strtrimdata, 5)
			fmtparcelnum = strleftChars + "-" + strmidChars +  "-" + strrightChars
		End If
	End If

If cid = 13  Then         'Creates  formatted parcel XX-XXXX-XXX  (2-4-3)
	If intLength = 1 Then
			fmtparcelnum = strtrimdata
		elseif intlength = 2 Then
			fmtparcelnum = strtrimdata
		elseif intlength = 3 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Right(strtrimdata, 1)
			fmtparcelnum = strleftChars + "-" + strmidChars
		elseif intlength = 4 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Right(strtrimdata, 2)
			fmtparcelnum = strleftChars + "-" + strmidChars
	'	Response.write(" the value of the fmtparcelnum in len 4 : " & fmtparcelnum)
		elseif intlength = 5 Then
			strleftChars = Left(strtrimdata, 2)
			strrightChars = Right(strtrimdata, 3)
			fmtparcelnum = strleftChars + "-" + strrightChars
	'	Response.write(" the value of the fmtparcelnum in len 5 : " & fmtparcelnum)
		elseif intlength = 6 Then
			strleftChars = Left(strtrimdata, 2)
			strrightChars = Right(strtrimdata, 4)
			fmtparcelnum = strleftChars + "-" + strrightChars
		'Response.write(" the value of the fmtparcelnum in len 6 : " & fmtparcelnum)
		elseif intLength = 7 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 4)
			strrightChars = Right(strtrimdata, 1)
			fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
			elseif intLength = 8 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 4)
			strrightChars = Right(strtrimdata, 2)
			fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
		elseif intlength = 9 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 4)
			strrightChars = Right(strtrimdata, 3)
			fmtparcelnum = strleftChars + "-" + strmidChars +  "-" + strrightChars
		End If
	End If

	Session("PrclNo") = fmtparcelnum
end Function
'**********
'The check for a * at the end of a search string. lem 3-12-09
'This function will be used in the Name, Address, Parcel, and Plat search field.
Function ChckStar(strSearch, checkstar)
	starwhere = InStr(strSearch, "*")
	strLen = Len(strSearch)
	'Response.Write(" the value of starwhere is : " & starwhere  )
	'Response.Write(" the value of strSearch is : " & strSearch  )

	if starwhere = 0 Then
	checkstar = strSearch + "*"
	'Response.Write(" the value of checkstar is : " & checkstar  )
	else
	'checkstar = strSearch + "*"
	checkstar = strSearch
	'Response.Write(" the value of checkstar is : " & checkstar  )
	strStarLoc = Mid(strSearch, strLen, 1)
	'Response.Write(" the value of the strStarLoc : " & strStarLoc )
		if strStarLoc = strLen then

		else
		' I have more than one * in the user filled form !
		End If

	End If
End Function


'*********
' The check for a hyphen function * * *    lem 9-16-08'
Function ChckHyphen(strNumber, checkit)
	'dim checkit as string
	'dim strNumber as string
	'dim strFirst as string'
	'Response.Write(" the value of strNumber is : " & strNumber  )
	strFirst = Trim(strFirst)
	strLen = Len(strNumber)
	strFirst = Left(strNumber, 1)

	if  strLen > 1 Then
	if  strFirst = "-" then
		checkit = 1
		exit function
	else
		checkit = 0
	end if
	End if
	if  strLen > 2 Then
	strFirst = Mid(strNumber, 2, 1)
		if  strFirst = "-" then
			checkit = 1
			exit function
		else
			checkit = 0
		end if

	End if
	if  strLen > 3 Then
	strFirst = Mid(strNumber, 3, 1)
		if  strFirst = "-" then
			checkit = 1
			exit function
		else
			checkit = 0
		end if
	End if
	if  strLen > 4 Then
	strFirst = Mid(strNumber, 4, 1)
		if  strFirst = "-" then
			checkit = 1
			exit function
		else
			checkit = 0
		end if
	End if
	if  strLen > 5 Then
	strFirst = Mid(strNumber, 5, 1)
		if  strFirst = "-" then
			checkit = 1
			exit function
		else
			checkit = 0
		end if
	End if
	if  strLen > 6 Then
	strFirst = Mid(strNumber, 6, 1)
		if  strFirst = "-" then
			checkit = 1
			exit function
		else
			checkit = 0
		end if
	End if
	if  strLen > 7 Then
	strFirst = Mid(strNumber, 7, 1)
		if  strFirst = "-" then
			checkit = 1
			exit function
		else
			checkit = 0
		end if
	End if
	if  strLen > 8 Then
	strFirst = Mid(strNumber, 8, 1)
		if  strFirst = "-" then
			checkit = 1
			exit function
		else
			checkit = 0
		end if

	End if
	if  strLen > 9 Then
	strFirst = Mid(strNumber, 9, 1)
		if  strFirst = "-" then
			checkit = 1
			exit function
		else
			checkit = 0
		end if

	End if
end Function

'*******  end the ChckHyphen Function here


Dim objCommand, objRS, strQueryString, strYEAR, strPID, strTID, objRS3, objRS5, objRSV, objRSPYGEN, objRSV2, objRSV3, objRS10, objRScount, intnumberQ
Dim intYYR, intYEAR

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE TXPRCL = '" & strQueryString & "'  AND TXFLAG='V' ;"

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
objCommand.CommandType = 1

Set objRSPYGEN = objCommand.Execute

Set objCommand = Nothing

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strQueryString & "' AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG='T';"'

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
objCommand.CommandType = 1

Set objRS = objCommand.Execute

Set objCommand = Nothing



Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strRID = request.QueryString("rid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
if strRID = "" or strRID = "0" then
strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' ORDER BY [Table 9 - Value Info].RecNum;"
else
strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' AND [Table 9 - Value Info].RecNum = " & strRID & " ;"
end if
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1
Set rs = objCommand.Execute

Set objCommand = Nothing


Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 3 - Part1 of RCPT Sets(1-5)].TXPRCL = '" & strQueryString & "'"

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 3 - Part1 of RCPT Sets(1-5)] " & strQueryString
objCommand.CommandType = 1

Set objRS3 = objCommand.Execute
Set objCommand = Nothing

Set objCommand = Server.CreateObject("ADODB.Command")
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 4 - Part2 of RCPT Sets(6-10)].TXPRCL = '" & strQueryString & "'"

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 4 - Part2 of RCPT Sets(6-10)] " & strQueryString
objCommand.CommandType = 1

Set objRS4 = objCommand.Execute
Set objCommand = Nothing



Set objCommand = Nothing

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 5 - Unpaid Taxes].TXPRCL = '" & strQueryString & "'"

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 5 - Unpaid Taxes] " & strQueryString
objCommand.CommandType = 1

Set objRS5 = objCommand.Execute
Set objCommand = Nothing
'CREATE THE OBJRSV FROM TABLE 9 **********************


Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "'"

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRSV = objCommand.Execute
intYEAR = objRSV("Year")

Set objCommand = Nothing


Set objCommand = Server.CreateObject("ADODB.Command")
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE (([Table 2 - Special/Ditch Info].TXPRCL) = '" & strQueryString & "') AND (([Table 2 - Special/Ditch Info].TXFLAG) ='V') "

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 2 - Special/Ditch Info] " & strQueryString
objCommand.CommandType = 1
Set objRSSP = objCommand.Execute
Set objCommand = Nothing




Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strRID = request.QueryString("rid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 9 - Value Info].PARCEL = '" & strQueryString & "' ORDER BY [Table 9 - Value Info].Parcel, [Table 9 - Value Info].Year DESC, [Table 9 - Value Info].RecNum DESC;"

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRSV2 = objCommand.Execute
Set objCommand = Nothing


Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strRID = request.QueryString("rid")
intYYR = request.QueryString("yr")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE (([Table 9 - Value Info].PARCEL) = '" & strQueryString & "') AND (([Table 9 - Value Info].RecNum = 1)) ORDER BY [Table 9 - Value Info].Parcel, [Table 9 - Value Info].Year DESC, [Table 9 - Value Info].RecNum DESC;"

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRSV3 = objCommand.Execute
Set objCommand = Nothing
' End of Create the OBJRSV3 from Table 9 * * * * * * * * * * * * * *



				strPID = request.QueryString("pid")
				intlenPID = len(strPID)
				if intlenPID = 12 then
				concstrPID = strPID
				end if
				if intlenPID = 11 then
				concstrPID = " " & strPID
				end if
				if intlenPID = 10 then
				concstrPID = " " & " " & strPID
				end if
				strTwn = Left(strPID, 2)
				p3prcl = Right(strPID, 4)
				p2prcla=Left(strPID, 6)
				p2prcl=Right(p2prcla, 3)

	Dim sesPID
	Session("sesPID") = concstrPID
	cesPID = Session("sesPID")



'* * * * * * *   Create the Record set for table 8     * * * *
Set objCommand = Server.CreateObject("ADODB.Command")

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 8 - Misc Data];"
objCommand.CommandType = 1

Set objRSCNTDT = objCommand.Execute
Set objCommand = Nothing


'CREATE THE OBJRSC10 FROM TABLE 10 **********************

					Set objCommand = Server.CreateObject("ADODB.Command")

					strPID = request.QueryString("pid")
					strTID = request.QueryString("tid")
					strRID = request.QueryString("rid")
					strQueryString = request.QueryString("pid")
					strQueryString = "WHERE (([Table 10 - City Special Asmt Info].PARCEL) = '" & strQueryString & "') AND (([Table 10 - City Special Asmt Info].[Rec#] = 0)) ORDER BY [Table 10 - City Special Asmt Info].Parcel;"

					objCommand.ActiveConnection = strConnect
					objCommand.CommandText = "SELECT * FROM [Table 10 - City Special Asmt Info] " & strQueryString
					objCommand.CommandType = 1

					Set objRSC10 = objCommand.Execute
					Set objCommand = Nothing
' End of Create the OBJRSC10 from Table 10 * * * * * * * * * * * * * *
'CREATE THE objRSCT10 FROM TABLE 10 **********************

					Set objCommand = Server.CreateObject("ADODB.Command")

					strPID = request.QueryString("pid")
					strTID = request.QueryString("tid")
					strRID = request.QueryString("rid")
					strQueryString = request.QueryString("pid")
					strQueryString = "WHERE (([Table 10 - City Special Asmt Info].PARCEL) = '" & strQueryString & "') AND (([Table 10 - City Special Asmt Info].[Rec#] > 0)) ORDER BY [Table 10 - City Special Asmt Info].Parcel, [Table 10 - City Special Asmt Info].[Rec#];"

					objCommand.ActiveConnection = strConnect
					objCommand.CommandText = "SELECT * FROM [Table 10 - City Special Asmt Info] " & strQueryString
					objCommand.CommandType = 1

					Set objRSCT10 = objCommand.Execute
					Set objCommand = Nothing
' End of Create the objRSCT10 from Table 10 * * * * * * * * * * * * * *



					Set objCommand = Server.CreateObject("ADODB.Command")

					strPID = request.QueryString("pid")
					strTID = request.QueryString("tid")
					strRID = request.QueryString("rid")

					strQueryString = request.QueryString("pid")
					strQueryString = "WHERE (([Table 11 - Value Special Asmt].TXPRCL) = '" & strQueryString & "') ORDER BY [Table 11 - Value Special Asmt].TXPRCL;"

					objCommand.ActiveConnection = strConnect
					objCommand.CommandText = "SELECT * FROM [Table 11 - Value Special Asmt] " & strQueryString
					objCommand.CommandType = 1

					Set objRS11 = objCommand.Execute
					Set objCommand = Nothing
' End of Create the objRS11 from Table 11 * * * * * * * * * * * * * *
%>

<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="280" >As of : <b><%=  objRSCNTDT("AsOfDate") %></b></td><td width="370" align="right" colspan="2">Parcel Number: <b><%= objRSPYGEN("TXPRCL") %></b></td>
	</tr>
	<tr valign="top">

</table>
<%

		If cid = 30 then
			Response.Write("<table border='0' cellpadding='0 cellspacing='0'>")
			Response.Write("<td  width='580'  align='right' >Value Year: </td>")
			Response.Write("<td width='74' align='right' class='tText3'>" & objRSPYGEN("TXYEAR") & "</td>")
		Else
			if objRSV2("RecNum") = 1 then
				Response.Write("<table border='0' cellpadding='0 cellspacing='0'>")
				Response.Write("<td width='650' align='right' colspan='2'>Payable Year: <b> " & objRSV2("Year") & " Rec# " & objRSV3("RecNum") & " of " & objRSV2("RecNum") & "</b></td>")
			else
				if strRID < objRSV2("RecNum") Then
					If strRID = "" then
					else
						Response.Write("<td width='650' align='right' colspan='2'>Payable Year: <b> " & objRSV2("Year") & " Rec# " & rs("RecNum") & " of " & objRSV2("RecNum") & "</b></td>")
					end if
				else
					if (rs("RecNum")) = (objRSV2("RecNum")) Then
						Response.Write("<table colspan='3' width='650' border='0' cellpadding='0 cellspacing='0'>")
						Response.Write("<td width='650' align='right' colspan='2'>Payable Year: <b> " & objRSV2("Year") & " Rec# " & rs("RecNum") & " of " & objRSV2("RecNum") & "</b></td>")
						if strRID = 0 then
							Response.Write("<table colspan='3' width='650' border='0' cellpadding='0 cellspacing='0'>")
							Response.Write("<td width='650' align='right' colspan='2'>Payable Year: <b>" & objRSV2("Year") & " Total </b></td>")
						end if
					else
						if strRID = 0 then
							Response.Write("<table colspan='3' width='650' border='0' cellpadding='0 cellspacing='0'>")
							Response.Write("<td  width='600'  align='right' >Payable Year: </td>")
							Response.Write("<td width='20' > <b>" & objRSV2("Year") & "</b></td>")
							Response.Write("<td width='80'  align='right'class='tText3'>Total Rec </td>")
						else
							Response.Write("<table colspan='3' width='650' border='0' cellpadding='0 cellspacing='0'>")
							Response.Write("<td width='650' align='right' colspan='2'>Payable Year: <b> " & objRSV2("Year") & " Rec# " & rs("RecNum") & " of " & objRSV2("RecNum") & "</b></td>")
						end if
					end if
				end if
			end if
		End if
%>

</table>
<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="650" align="right" colspan="2"> <b>
		<%

		if objRSV2("RecNum") = 1 then
		else
			if strRID < objRSV2("RecNum") Then
				If strRID = "" then
					strRID = 0
					Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID + 1 &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>Next Record2</a>")
				else
					Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID + 1 &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>Next Record3</a>")
					if strRID >= 1 Then
						Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID - 1 &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>| Previous Record</a>")
						Response.Write(" the value for the objRSV2(RecNum) :  " & objRSV2("RecNum") )
					end if
				end if

				if strRID = rid then
				else
					rs.MoveNext
				end if
			else
				if (rs("RecNum")) = (objRSV2("RecNum")) Then
					Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid=0&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>Total </a>")
					Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID - 1 &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>| Previous Record</a>")
				else
					if strRID >= 1 Then
						Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID + 1 &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>Next Record</a>")
						Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID - 1 &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>| Previous Record</a>")
					end if
					if strRID = 0 then
						Response.Write("<a class='pLink' href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid=1&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>| Detail Record</a>")
					end if
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
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=0&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>General Information</a>   |   ")
				If cid <> 30 then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Value Information</a>   |   ")
				End if
				If cid<> 30 or (cid = 30 and strTwn = 65) then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=2&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				End if
				If Not objRS.EOF then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=3&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>   ")
				End if
				If (cid = 47 and strTwn = 74) then
					response.Write(" | <a href='/tax/Data Sheet/" & strTwn & " "& p2prcl &" "& p3prcl &".pdf' target='_blank' class='uLink'>Data Sheet</a>   ")
				End if
			Case 1
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=0&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Information</a>    |   ")
				If cid <> 30 then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>Value Information</a>   |   ")
				End if
				If cid<> 30 or (cid = 30 and strTwn = 65) then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=2&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				End if
				If Not objRS.EOF then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=3&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>   ")
				End if
				If (cid = 47 and strTwn = 74) then
					response.Write(" | <a href='/tax/Data Sheet/" & strTwn & " "& p2prcl &" "& p3prcl &".pdf' target='_blank' class='uLink'>Data Sheet</a>   ")
				End if
			Case 2
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=0&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Information</a>   |   ")
				If cid <> 30 then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Value Information</a>   |   ")
				End if
				If cid<> 30 or (cid = 30 and strTwn = 65) then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=2&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>Special Asmts</a>   |   ")
				End if
				If Not objRS.EOF then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=3&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>   ")
				End if
				If (cid = 47 and strTwn = 74) then
					response.Write(" | <a href='/tax/Data Sheet/" & strTwn & " "& p2prcl &" "& p3prcl &".pdf' target='_blank' class='uLink'>Data Sheet</a>   ")
				End if
			Case 3
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=0&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Information</a>   |   ")
				If cid <> 30 then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Value Information</a>   |   ")
				End if
				If cid<> 30 or (cid = 30 and strTwn = 65) then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=2&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				End if
				If Not objRS.EOF then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=3&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>History</a>   ")
				End if
				If (cid = 47 and strTwn = 74) then
					response.Write(" | <a href='/tax/Data Sheet/" & strTwn & " "& p2prcl &" "& p3prcl &".pdf' target='_blank' class='uLink'>Data Sheet</a>   ")
				End if
			Case 4
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=0&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Information</a>   |   ")
				If cid <> 30 then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Value Information</a>   |   ")
				End if
				If cid<> 30 or (cid = 30 and strTwn = 65) then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=2&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |    ")
				End if
				If Not objRS.EOF then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=3&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>")
				End if
				If (cid = 47 and strTwn = 74) then
					response.Write(" | <a href='/tax/Data Sheet/" & strTwn & " "& p2prcl &" "& p3prcl &".pdf' target='_blank' class='uLink'>Data Sheet</a>   ")
				End if
			Case 5
				response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=0&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Information</a>   |   ")
				If cid <> 30 then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=1&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Value Information</a>   |   ")
				End if
				If cid<> 30 or (cid = 30 and strTwn = 65) then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=2&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				End if
				If Not objRS.EOF then
					response.Write("<a href='Parcel_value.asp?pid=" & strPID & "&tid=3&rid="& strRID &"&yr="& objRSPYGEN("TXYEAR") &"&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>History</a>   ")
				End if
				If (cid = 47 and strTwn = 74) then
					response.Write(" | <a href='/tax/Data Sheet/" & strTwn & " "& p2prcl &" "& p3prcl &".pdf' target='_blank' class='uLink'>Data Sheet</a>   ")
				End if
			End Select
		%>
		</td>
	</tr>
	<tr>
		<td width="650" bgcolor="#000000" height="1" colspan="2"></td>
	</tr>

	<tr valign="top">
		<td height="10"></td>
	</tr>
	<tr valign="top">
	<%

		if objRSV3("Text1") ="" and  objRSV3("Text2") ="" and objRSV3("Text3") ="" and objRSV3("Text4") ="" and objRSV3("Text5") =""  then
		Response.Write("<td width='450' align=left class='tlink2' colspan='2'> " & objRSV3("Text1") & "  " & objRSV3("Text2") & "   " & objRSV3("Text3") & " , " & objRSV3("Text4") & " , " & objRSV3("Text5") & "</td>")
		end if
		if didit <> 3 then
		if objRSV3("Text5") <> ""  then ' and  objRSV3("Text2") =""  and objRSV3("Text3") ="" and objRSV3("Text4") ="" and objRSV3("Text5") =""  then
		Response.Write("<td width='450' align=left class='tlink2' colspan='2'> " & objRSV3("Text1") & " , " & objRSV3("Text2") & "  , " & objRSV3("Text3") & " , " & objRSV3("Text4") & " , " & objRSV3("Text5") & "</td>")
		didit = 3
		end if
		end if
		if didit <> 3 then
		if objRSV3("Text4") <> ""  then ' and  objRSV3("Text2") ="" then ' and objRSV3("Text3") ="" and objRSV3("Text4") ="" and objRSV3("Text5") =""  then
		Response.Write("<td width='450' align=left class='tlink2' colspan='2'> " & objRSV3("Text1") & " , " & objRSV3("Text2") & "  , " & objRSV3("Text3") & " , " & objRSV3("Text4") & "  " & objRSV3("Text5") & "</td>")
		didit = 3
		end if
		end if
		if didit <> 3 then
		if objRSV3("Text2") <> ""  then 'and  objRSV3("Text2") <>"" and objRSV3("Text3") <>"" then ' and objRSV3("Text4") ="" then ' and objRSV3("Text5") =""  then
		Response.Write("<td width='450' align=left class='tlink2' colspan='2'> " & objRSV3("Text1") & " , " & objRSV3("Text2") & "   " & objRSV3("Text3") & "  " & objRSV3("Text4") & "  " & objRSV3("Text5") & "</td>")
		didit = 3
		end if
		end if
		if didit <> 3 then
		if objRSV3("Text1") <> ""  then 'and  objRSV3("Text2") <>"" and objRSV3("Text3") <>"" and objRSV3("Text4") <>"" then ' and objRSV3("Text5") =""  then
		Response.Write("<td width='450' align=left class='tlink2' colspan='2'> " & objRSV3("Text1") & "  " & objRSV3("Text2") & "   " & objRSV3("Text3") & "  " & objRSV3("Text4") & "  " & objRSV3("Text5") & "</td>")
		didit = 3
		end if
		end if

	%>
	</tr>

	<tr valign="top">
		<td width="10"></td>
		<td width="650">
		<%
			Select Case strTID
			Case 0
		%>
		<!-- #include file="GeneralTax_value.asp" -->
		<%
			Case 1
		%>
		<!-- #include file="ValueInformation.asp" -->
		<%
			Case 2
			If cid = 30 and strTwn = 65 then
		%>
		<!-- #include file="SpecialAssessment_Mandan.asp" -->
		<%
			Else
			if cid = 30 then
		%>
		<!-- #include file="SpecialAssessment.asp" -->

		<%
			end if
			end if
			If cid = 47 then
		%>
		<!-- #include file="SpecialAssessment.asp" -->

		<%
			end if
		%>
		<%
			Case 3
		   If Not objRS.EOF then
		%>
		<!-- #include file="History.asp" -->
		<%
			end if
			End Select
		%>
		</td>
	</tr>
	</table>
	<table border="0" cellpadding="0" cellspacing="0">
	<tr>
	<td width="10"></td>
	<td width="650" align="Right" >
	<tr>
	<td height="15" colspan="2"></td>
	</tr>
	<td width="1000" align="right" class="STitle2">
<%
			Response.Write("<a href='searchinput.asp?cid=" & cid & "' class='tlink'>Another Search    |</a>&nbsp;&nbsp;")
			Response.Write("<a href='ParcelListReturn.asp?pid=" & strPID & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange  & "&varstrLot=" & strLot & "' class='tlink'>Back to ParcelList    |</a>&nbsp;&nbsp")
%>
			</td>
	</tr>
		<td height="15"></td>
	</tr>
	</table>


	<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
<%
	if cid = 23 then
%>
		<td width="180"  class="STitle1">
		Auditor<br>
		Michial Johnson<br>
		PO Box 128<br>
		Lamoure ND 58458-0128<br>
			</td>
		<td width="200"  class="STitle1">
		Treasurer<br>
		Lamoure County Treasurer<br>
		PO Box 128<br>
		Lamoure ND 58458-0128<br>
		701-883-5301 Ext 216
			</td>
<%
	end if
%>
<%
	if cid = 37 then
%>
		<td width="180"  class="STitle1">
		Auditor<br>
		Connie Gilbert<br>
		PO Box 668<br>
		Lisbon ND 58054-0668<br>
			</td>
		<td width="200"  class="STitle1">
		Treasurer<br>
		Ransom County Treasurer<br>
		PO Box 629<br>
		Lisbon ND 58054-0629<br>
		</td>
<%
	end if
%>
<%
	if cid = 41 then
%>
		<td width="180"  class="STitle1">
		Auditor<br>
		Sherry Hosford<br>
		355 Main St S, Suite 1<br>
		Forman ND 58032-4149<br>
			</td>
		<td width="200"  class="STitle1">
		Treasurer<br>
		Gina Hillestad<br>
		355 Main St S, Suite 4<br>
		Forman ND 58032-4149<br>
		</td>
<%
	end if
%>
<%
	if cid = 47 then
%>
		<td width="180"  class="STitle1">
		Auditor<br>
		Casey Bradley<br>
		511 2nd Ave SE, Ste 102<br>
		Jamestown ND 58401<br>
			</td>
	 	<td width="200"  class="STitle1">
		Treasurer<br>
		(701-252-9036)<br>
		511 2nd Ave SE, Ste 101<br>
		Jamestown ND 58401<br>
		</td>
<%
	end if
%>

		</td>
</table>
<%
objRS.Close
Set objRS = Nothing
%>
</body>
</html>