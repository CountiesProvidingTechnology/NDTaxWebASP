<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Parcel Search Results</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
	Dim cid, cName, sesTotal
	'cid = Session("CountyID")
	varcid = request.QueryString("cid")
	intParcelNo = request.QueryString("varintParcelNo")
	'response.write(" the value of intParcelNo : " & intParcelNo )
	strAddress = request.QueryString("varstrAddress")
	strName = request.QueryString("varstrName")
	intSect = request.QueryString("varintSect")
	intTwp = request.QueryString("varintTwp")
	intRange = request.QueryString("varintRange")


	Session("sesPID") = concstrPID
	'cName = Session("County Name")
	sesTotal = Session("amtTotal")
	tab4 = Session("recnumberend")
	Dim strCID
	strCID = "WebTab" & Session("CountyID")
	'response.Write (" Session Var " & cid  )
	'response.Write(" the value of strCID::" & strCID )
	'response.write(" the new varcid :" & varcid )
	response.Write("<link rel='stylesheet' href='" & varcid & ".css' type='text/css'>")


	'Get the county name based on the county number
		Select Case varcid
				Case 21
				cName = "Douglas"
				Case 26
				cName = "Grant"
				Case 34
				cName = "Kandiyohi"
				Case 41
				cName = "Lincoln"
				Case 42
				cName = "Lyon"
				Case 45
				cName = "Marshall"
				Case 47
				cName = "Meeker"
				Case 48
				cName = "Mille Lacs"
				Case 51
				cName = "Murray"
				Case 53
				cName = "Nobles"
				Case 54
				cName = "Norman"
				Case 61
				cName = "Pope"
				Case 64
				cName = "Redwood"
				Case 65
				cName = "Renville"
				Case 67
				cName = "Rock"
				Case 74
				cName = "Steele"
				Case 75
				cName = "Stevens"
				Case 76
				cName = "Swift"
				Case 77
				cName = "Todd"
				Case 78
				cName = "Traverse"
				Case 80
				cName = "Wadena"
				Case 84
				cName = "Wilkin"
				Case 87
				cName = "Yellow Medicine"
		End Select


%>
<script type="text/javascript">

//tmpcid = <%=Session("CountyID")%>;

//strPID = <%=request.QueryString("pid")%>;

function open_win(tmpCID, tmpPID)
{
	alert(" made it");
	//var tmpPI = tmpPID + "";
	//var tmpPID = <%=request.QueryString("pid")%> + "";
	//alert(tmpPI);
	alert(tmpCID);
	alert(tmpPID);
window.open("http://morris.state.mn.us:41080/iText/txstmt.jsp?cid=tmpCID&pid=tmpPID");
}
</script>


</head>
<!-- #include file="insDB.asp" -->


<body>
<%
Function printTaxRecord()
	Dim recordNumber
	'Dim checkit as boolean
	'Dim strNumber as string
	For recordNumber=1 to 5
	Session("amtTotal") = 0
	sesTotal = 0
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
				'Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCSAA" & recordNumber & "1")), 2) & recordNumber &"</td>")'
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCSAA" & "1" & recordNumber)),2,0,vbTrue,0) & "</td>")
				'Response.Write("<td width='80' class='rText' align='right'>" & objRS3("RCSAA" & "1" & recordNumber) & "</td>")
				sesTotal = sesTotal + objRS3("RCSAA" & "1" & recordNumber)
				Session("recnumberend") = recordNumber
				'tab4 = Session("recnumberend")
				Response.Write("</tr>")
			printTaxValues(recordNumber)
		Session("amtTotal") = 0
		sesTotal = 0
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
				'the line below will add a new cell at the end of an existing cell
				' this will add the recordnumber on the outside of the main part of the page.
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
					'Response.Write("<td width='80' class='rText'>test "& subrecordnumber & recordNumber & "</td>")
				If (objRS3(("RCSAA2" & recordNumber )) = 0) or (objRS3(("RCSAA3" & recordNumber )) = 0) Then
					'Response.Write("<td width='80' class='rText'>RCSAA2/3=0</td>")
						If (objRS3(("RCSAA41")) = 0)  Then
							'Response.Write("<td width='80' class='rText'>RCSAA41=0</td>")
							'Response.Write("<td width='80' class='rText'></td>")
						else
							'Response.Write("<td width='80' class='rText'>RCSAA41=0ELSE</td>")
						End if
						If (objRS3(("RCSAA42")) = 0)  Then
							'Response.Write("<td width='80' class='rText'>RCSAA42=0</td>")
							'Response.Write("<td width='80' class='rText'></td>")
						else
							'Response.Write("<td width='80' class='rText'>RCSAA42=0ELSE</td>")
							'Response.Write("<td width='80' class='rText'></td>")
						End if
				else
					'Response.Write("<td width='80' class='rText'>RCSAA2/3=0ELSE</td>")
					Response.Write("<td width='80' class='rText'>&nbsp;</td>")
				End if
				If (objRS3(("RCSAA4" & recordNumber )) = 0)  Then
					'Response.Write("<td width='80' class='rText'>RCSAA4=0</td>")
				else

					'Response.Write("<td width='80' class='rText'>RCSAA4=0ELSE</td>")
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
						If recordNumber < 5 Then
							if (objRS3(("RCAMT2" & recordNumber + 1 )) = 0) or (objRS3(("RCSAA3" & recordNumber + 1 )) = 0) Then
								'Response.Write("<td width='80' class='rText'>RCSAA3+1" & subrecordnumber & recordnumber & "</td>")
								Response.Write("<td width='80' class='rText'>&nbsp;</td>")
							else
								'Response.Write("<td width='80' class='rText'>RCSAA3+1" & subrecordnumber & recordnumber & "</td>")
								Response.Write("<td width='80' class='rText'>&nbsp;</td>")
							end if
						else
							Response.Write("<td width='80' class='rText'>&nbsp;</td>")
						end if
					end if
						Response.Write("</tr>")
				tempsubrecordnumber = subrecordnumber
				subrecordNumber = 7
		End if
		Next
		'Response.Write("<td width='80' class='rText'>"& subrecordnumber & recordNumber & "</td>")
	'end if
	End Function

Function printTaxRecord4()
	Dim recordNumber
	For recordNumber=6 to 10
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
			if recordnumber < 10 then
			If objRS4("RCBTDT" & (recordNumber + 1)) = 0 Then
			'Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")'
			else
				'Response.Write("<td width='80' class='rText' align='right'>" & recordnumber & subrecordnumber & "</td>")
			end if
			end if
			'Response.Write("</tr>")
			recordNumber = 10
		Else
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='260' class='rText' align='left' rowspan='6'>")
				'Response.Write("<b>" & calcDateRS4(("RCBTDT" & recordNumber), (recordNumber)) & "</b><br>")'
				Response.Write("<b>" & calcDateRS4(("RCBTDT"), recordnumber) & "</b><br>")
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
				Session("recnumberend") = recordNumber
				sesTotal = sesTotal + objRS4("RCSAA" & "1" & recordNumber)
				Response.Write("</tr>")
			printTaxValues4(recordNumber)
		End If
	Next
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

Function calcZeroT(strData, intPlaces)
	If objRST(strData) <> "" Then
		strData = (FormatNumber(objRST(strData), intPlaces))
	Else
		strData = "0.00"
	End If
	calcZeroT = strData
end Function

Function calcZeroRS5(strData, intPlaces)
	If objRS5(strData) <> "" Then
		strData = (FormatNumber(objRS5(strData), intPlaces))
	Else
		strData = "0.00"
	End If
	calcZeroRS5 = strData
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

Function calcZip(strData)
	If objRS(strData) = "00000" Then
		strData = ""
	End If
	calcZip = strData
end Function

'*********
' The negative number function to check if a value is negative and to rewrite it as a string for display
'Function Negnum(strNumber, checkit)
'	dim checkit as boolean
'	dim strNumber as string
'	dim strFirst as string
'	strFirst = Left(strNumber, 1)
'	if  strFirst = "-" then
'		checkit = 1
'		Negnum = strNumber
'	else
'		checkit = 0
'		Negnum = strNumber
'	end if
'
'end Function

'*******  end the Negnum Function here


Dim objCommand, objRS, strQueryString, strPID, strTID, objRS3, objRS5, objRScount, intnumberQ, objRST

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

'****      for the recordset that has the 'T' value in the  FLAG field. * * * * * * * *
'Dim objRST, strQueryString, strPID, strTID
Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strQueryString & "' AND (([Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG) ='T')"
'Response.Write("The value of the Query String : " & strQueryString )

'Fill in the command properties
'Response.Write( "the strConnect is " & strConnect   )
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
objCommand.CommandType = 1

Set objRST = objCommand.Execute

Set objCommand = Nothing
'******      end the recordset for the 'T'   value in the FLAG field of Table 1   * * * * * * * * * * *

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

'*****   to create the RS4  recordset
Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 4 - Part2 of RCPT Sets(6-10)].TXPRCL = '" & strQueryString & "'"

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 4 - Part2 of RCPT Sets(6-10)] " & strQueryString
objCommand.CommandType = 1

Set objRS4 = objCommand.Execute
Set objCommand = Nothing

'****     the end of the RS4 recordset

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

'*****   to create the RS6  recordset
Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 6 - Part3 of RCPT Sets(11-15)].TXPRCL = '" & strQueryString & "'"

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 6 - Part3 of RCPT Sets(11-15)] " & strQueryString
objCommand.CommandType = 1

Set objRS6 = objCommand.Execute
Set objCommand = Nothing

'****     the end of the RS6 recordset

Set objCommand = Nothing

'CREATE THE OBJRSSPT *****************
Set objCommand = Server.CreateObject("ADODB.Command")
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = request.QueryString("yr")
intYYR = Session("intYEAR")
'Response.Write(" value of the intYYR :" & intYYR )
strQueryString = request.QueryString("pid")

strQueryString = "WHERE ((([Table 2 - Special/Ditch Info].TXPRCL) = '" & strQueryString & "') AND (([Table 2 - Special/Ditch Info].TXFLAG) ='T')  AND  (([Table 2 - Special/Ditch Info].TSSAC1) >0))"
'strQueryString = "WHERE ((([Table 2 - Special/Ditch Info].TXPRCL) = '" & strQueryString & "') AND (( [Table 2 - Special/Ditch Info].TXYEAR) =" & intYYR & ")  AND  (([Table 2 - Special/Ditch Info].TXFLAG) ='T')  AND  (([Table 2 - Special/Ditch Info].TSSAR1) >0)) "

'strQueryString = "WHERE ((([Table 2 - Special/Ditch Info].TXPRCL) = '" & strQueryString & "') AND (( [Table 2 - Special/Ditch Info].TXYEAR) =" & intYYR & ")  AND  (([Table 2 - Special/Ditch Info].TXFLAG) ='V')) "
'strQueryString = "WHERE [Table 2 - Special/Ditch Info].TXPRCL = '" & strQueryString & "' AND [Table 2 - Special/Ditch Info].TXYEAR =" & intYYR & " "

'strQueryString = "WHERE [Table 2 - Special/Ditch Info Query].TXPRCL = '" & strQueryString & "' AND [Table 2 - Special/Ditch Info Query].TXYEAR =intYEAR; "
'strQueryString = "WHERE [Table 2 - Special/Ditch Info Query].TXPRCL = '" & strQueryString & "' AND [Table 2 - Special/Ditch Info Query].TXYEAR =[intYEAR];"
'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 2 - Special/Ditch Info] " & strQueryString
'objCommand.CommandText = "SELECT * FROM [Table 2 - Special/Ditch Info Query] " & strQueryString
objCommand.CommandType = 1
'Response.Write(" the value of strQueryString : " & strQueryString )
Set objRSSPT = objCommand.Execute
Set objCommand = Nothing


'* * * * * * *   Create the Record set for table 8     * * * *
'This will get the value for the date in the "As of"   date

'Dim objCommand, objRSCNT, strQueryString, strPID, strTID, objRS3, objRS5, objRScount, intnumberQ

Set objCommand = Server.CreateObject("ADODB.Command")

'strPID = request.QueryString("pid")
'strTID = request.QueryString("tid")
'strQueryString = request.QueryString("pid")
'strQueryString = "WHERE [Table 8 - Misc Data].Counter = '" & strQueryString & "'"
'Response.Write("The value of the Query String : " & strQueryString )

'Fill in the command properties
'Response.Write( "the strConnect is " & strConnect   )
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 8 - Misc Data];"
objCommand.CommandType = 1
'
Set objRSCNTDT = objCommand.Execute
'
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
' End of Create the OBJRSV2 from Table 9 * * * * * * * * * * * * * *
'CREATE THE OBJRSV3 FROM TABLE 9 **********************

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strRID = request.QueryString("rid")
intYYR = request.QueryString("yr")
'response.write("the value for intYYR : " & intYYR )
strQueryString = request.QueryString("pid")
strQueryString = "WHERE (([Table 9 - Value Info].PARCEL) = '" & strQueryString & "') AND (([Table 9 - Value Info].RecNum = 1)) ORDER BY [Table 9 - Value Info].Parcel, [Table 9 - Value Info].Year DESC, [Table 9 - Value Info].RecNum DESC;"

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRSV3 = objCommand.Execute
Set objCommand = Nothing
' End of Create the OBJRSV3 from Table 9 * * * * * * * * * * * * * *

'CREATE THE OBJRS11 FROM TABLE 11 ********************** Created 9-10-08  L E M

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")


intYYR = request.QueryString("yr")
'response.write("the value for intYYR : " & intYYR )
strQueryString = request.QueryString("pid")
'strQueryString = "WHERE (([Table 11 - Legal Descriptions(9+)].TXPRCL) = '" & strQueryString & "') AND (([Table 11 - Legal Descriptions(9+)].TXYEAR = '" & intYYR & "')) ORDER BY [Table 11 - Legal Descriptions(9+)].TXPRCL, [Table 11 - Legal Descriptions(9+)].TXYEAR, [Table 11 - Legal Descriptions(9+)].TXREC;"
strQueryString = "WHERE (([Table 11 - Legal Descriptions(9+)].TXPRCL) = '" & strQueryString & "')  AND (([Table 11 - Legal Descriptions(9+)].TXFLAG = 'T')) ORDER BY [Table 11 - Legal Descriptions(9+)].TXPRCL, [Table 11 - Legal Descriptions(9+)].TXYEAR, [Table 11 - Legal Descriptions(9+)].TXREC;"

'Fill in the command properties
objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 11 - Legal Descriptions(9+)] " & strQueryString
objCommand.CommandType = 1

Set objRS11 = objCommand.Execute
Set objCommand = Nothing
' End of Create the OBJRS11 from Table 11 * * * * * * * * * * * * * *



'********    the calculation of the length of the strPID variable and the correction of the
' length for the SELECT statement in the query for the record set below.
'3-13-08 lem it seems that the TXSTMT XX   table has no spaces ahead of the parcel number.
'  this is an imported table and therefore it not to be considered 'the same ' as the other
' tables in the DATABASE.
'response.Write(" the value of cid : " & cid & " and the value of strPID : " & strPID )
'   code here to get the length of the parcel variable. Then add a blank to the front of the
' variable and create a new variable for the HREF line.  3-12-08
				intlenPID = len(strPID)
				'response.Write(" the value of intlenPID : " & intlenPID & strPID )
				if intlenPID = 12 then
				concstrPID = strPID
				'response.Write(" the value of concstrPID : " & concstrPID & strPID )
				end if
				if intlenPID = 11 then
				concstrPID = " " & strPID
				'response.Write(" the value of concstrPID : " & concstrPID & concstrPID )
				end if
				if intlenPID = 10 then
				concstrPID = " " & " " & strPID
				end if
	Dim sesPID
	Session("sesPID") = concstrPID
	cesPID = Session("sesPID")
'*******      the calculation of the  length of the strPID variable is done above. * * * * * * *


'CREATE THE OBJRSTXST FROM TABLE TXSTMT XX **********************

'Set objCommand = Server.CreateObject("ADODB.Command")
'
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strRID = request.QueryString("rid")
intYYR = request.QueryString("yr")
'strQueryString = concstrPID
'strQueryString = "WHERE (([TXSTMT"& varcid &"].FIELD2) ='" & strQueryString & "')"

'Fill in the command properties
'objCommand.ActiveConnection = strConnect
'objCommand.CommandText = "SELECT * FROM [TXSTMT"& varcid &"] " & strQueryString
'objCommand.CommandType = 1
'Set objRSTXST = objCommand.Execute
'Set objCommand = Nothing
' End of Create the OBJRSTXST from Table TXSTMT XX * * * * * * * * * * * * * *
intYrs = Session("intTAXYEAR")


%>

<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
<td width="120" >As of : <b><%=  objRSCNTDT("AsOfDate") %></b></td><td width="510" align="right" colspan="2">Parcel Number: <b><%= objRS("TXPRCL") %></b></td>
	</tr>
	<tr valign="top">
		<td width="650" align="right" colspan="2">Payable Year: <b><%= objRST("TXYEAR") %></b></td>
<%
			Select Case strTID
			Case 0
'  ****  this line will handle the older tax statement for the user to print off of the web site
'response.Write("<td width='400' align='right' colspan='2'><A HREF='http://morris.state.mn.us:41080/iText/TXSTMT2009.jsp?cid=" & varcid & "&pid=" & concstrPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
response.Write("<td width='400'></td><td><A HREF='http://morris.state.mn.us:41080/iText/TXSTMT2010.jsp?cid=" & varcid & "&pid=" & concstrPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
			end Select

%>

	</tr>


<%


%>
</table>
<table border="0" cellpadding="0" cellspacing="0">

	<tr>
		<td width="10"></td>
		<td width="1000" align="left" colspan="6">
		<%
		myname=objRS("TXTNAM")
		myname=Replace(myname,"&","%26")
		Location= Instr(myname,"/")
 		if Location>0 then
		lastname=Left(myname,Location-1)
		End if
		
		length=Len(myname)
		firstname=Right(myname,length-Location)
		
			Select Case strTID
			Case 0
				
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Unpaid Tax</a>   |    ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>   ")
					If cName = "Swift" then
					Response.Write("<a href='https://swp.paymentsgateway.net/default.aspx?pg_api_login_id=X6p8j5PGd0&pg_consumerorderid=" & strPID & "&pg_billto_postal_name_first=" & firstname &"&pg_billto_postal_name_last=" & lastname &"' ) class='tLink2' align='right'>          E-payment    </a> |  ")
					'response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=222859866205626577232622909126840'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?clientId=222859866205626577232605729257656&cde-ParcNumb-0=" & strPID & "'  class='tlink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
					end if
					If cName = "Renville" then
					'Response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Kandiyohi" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=161738092023567529211240682237982904&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Lyon" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764220341208479929&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Pope" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card  </a> ")
					Response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>    Pay By echeck</a> ")
					end if
					If cName = "Redwood" then
					Response.Write("<a href='http://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;          Pay By Credit Card    </a> |  ")
					Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  Pay By echeck</a> ")
					'Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>  Pay By echeck</a> ")
					end if
					If cName = "Yellow Medicine" then
					'Response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=34971905467269240275222972620077240&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=34971905467269240275222972620077240&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Steele" then

					'response.Write("<a href='https://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=3338095734343483515010086874118329'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?clientId=3338095734343483515009979499935929'  class='ulink' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Credit Card Payment</a> ")
					response.Write("|<a href='http://www.co.steele.mn.us/Property_Data_Tax_Search.html' class='uLink'>Map </a>")
					response.Write("<a href='http://www.co.steele.mn.us/TREASUR/RE_Tax_Pmt_Options.html'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes</a> ")
					end if

				If cName = "Douglas" Then
				 response.Write("|<a href='http://206.145.187.195/douglas_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'>View Maps</a>")
				 Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447712482162873&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By Credit Card</a> ")
				end if
				If cName = "Rock" Then
				response.write("|<a href='http://rock.houstoneng.com/?call=search_parcels_taxnum&value0=" & strPID & "'class='ulink'> View Maps </a>")
				end if


'' **** line below is the link to print the current year tax statement on this page. ***   LEM

response.Write("<td><A HREF='http://morris.state.mn.us:41080/iText/TXSTMT2011.jsp?cid=" & varcid & "&pid=" & concstrPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")

			Case 1
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Unpaid Tax</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>   ")

				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='hLink'>Tax Information</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>   ")
					If cName = "Swift" then
					Response.Write("<a href='https://swp.paymentsgateway.net/default.aspx?pg_api_login_id=X6p8j5PGd0&pg_consumerorderid=" & strPID & "&pg_billto_postal_name_first=" & firstname &"&pg_billto_postal_name_last=" & lastname &"' ) class='tLink2' align='right'>          E-payment    </a> |  ")
					'response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=222859866205626577232622909126840'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=222859866205626577232622909126840&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
					end if
					If cName = "Renville" then
					'Response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Kandiyohi" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=161738092023567529211240682237982904&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Lyon" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764220341208479929&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Pope" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
					end if
					If cName = "Redwood" then
					Response.Write("<a href='http://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>  Pay By Credit Card</a> ")
					Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>  Pay By echeck</a> ")
					'Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>   Pay By echeck</a> ")
					end if
					If cName = "Yellow Medicine" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=34971905467269240275222972620077240&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Steele" then

					'response.Write("<a href='https://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=3338095734343483515010086874118329'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?clientId=3338095734343483515009979499935929'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes</a> ")
					response.Write("|<a href='http://www.co.steele.mn.us/Property_Data_Tax_Search.html' class='uLink'>Map </a>")
					response.Write("<a href='http://www.co.steele.mn.us/TREASUR/RE_Tax_Pmt_Options.html'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes</a> ")
					end if

				If cName = "Douglas" Then
				 response.Write("|<a href='http://206.145.187.195/douglas_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> View Maps </a>")
				 Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447712482162873&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By Credit Card</a> ")
				end if
				If cName = "Rock" Then
				response.write("|<a href='http://rock.houstoneng.com/rock/rock.html?call=ItemQuery&buffer=500&layers=Parcels&qstring=" & strPID & "'class='ulink'> View Maps </a>")
				end if


			Case 2
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Unpaid Tax</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>   ")

				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='hLink'>Current Receipts</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>   ")
					If cName = "Swift" then
					Response.Write("<a href='https://swp.paymentsgateway.net/default.aspx?pg_api_login_id=X6p8j5PGd0&pg_consumerorderid=" & strPID & "&pg_billto_postal_name_first=" & firstname &"&pg_billto_postal_name_last=" & lastname &"' ) class='tLink2' align='right'>          E-payment    </a> |  ")
					'response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=222859866205626577232622909126840'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?clientId=222859866205626577232605729257656' class='tlink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     Pay By Credit Card</a> ")
					'response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=222859866205626577232622909126840&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
					end if
					If cName = "Renville" then
					'Response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Kandiyohi" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=161738092023567529211240682237982904&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Lyon" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764220341208479929&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Pope" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
					end if
					If cName = "Redwood" then
					Response.Write("<a href='http://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>  Pay By Credit Card</a> ")
					Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>  Pay By echeck</a> ")
					'Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>   Pay By echeck</a> ")
					end if
					If cName = "Yellow Medicine" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=34971905467269240275222972620077240&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Steele" then

					'response.Write("<a href='https://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=3338095734343483515010086874118329'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?clientId=3338095734343483515009979499935929'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes</a> ")
					response.Write("|<a href='http://www.co.steele.mn.us/Property_Data_Tax_Search.html' class='uLink'>Map </a>")
					response.Write("<a href='http://www.co.steele.mn.us/TREASUR/RE_Tax_Pmt_Options.html'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes</a> ")

					end if

				If cName = "Douglas" Then
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|   |")
				'	response.Write("<a href='http://morris.state.mn.us/dcmap/' class='ulink'>View Maps </a>")
				'	'response.Write("<a href='javascript:zoomto_selectfeature('strPIN');' class='uLink'> View Maps </a>")
				'   response.Write("<a href='http://ipaddress/douglas_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> Maps </a>")
				 response.Write("|<a href='http://206.145.187.195/douglas_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> View Maps </a>")
				 Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447712482162873&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By Credit Card</a> ")
				'else
				'	response.Write("|<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				end if
				If cName = "Rock" Then
				' response.Write("|<a href='http://206.145.187.195/rock_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> View Maps </a>")
				response.write("|<a href='http://rock.houstoneng.com/rock/rock.html?call=ItemQuery&buffer=500&layers=Parcels&qstring=" & strPID & "'class='ulink'> View Maps </a>")
				end if

			Case 3
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Unpaid Tax</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>   ")

				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='hLink'>Unpaid Tax</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>   ")
					If cName = "Swift" then
					Response.Write("<a href='https://swp.paymentsgateway.net/default.aspx?pg_api_login_id=X6p8j5PGd0&pg_consumerorderid=" & strPID & "&pg_billto_postal_name_first=" & firstname &"&pg_billto_postal_name_last=" & lastname &"' ) class='tLink2' align='right'>          E-payment    </a> |  ")
					'response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=222859866205626577232622909126840'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?clientId=222859866205626577232605729257656' class='tlink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     Pay By Credit Card</a> ")
					'response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=222859866205626577232622909126840&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
					end if
					If cName = "Renville" then
					'Response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Kandiyohi" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=161738092023567529211240682237982904&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Lyon" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764220341208479929&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Pope" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
					end if
					If cName = "Redwood" then
					Response.Write("<a href='http://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>   Pay By Credit Card</a> ")
					Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>  Pay By echeck</a> ")
					'Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>  Pay By echeck</a> ")
					end if
					If cName = "Yellow Medicine" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=34971905467269240275222972620077240&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Steele" then

					'response.Write("<a href='https://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=3338095734343483515010086874118329'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?clientId=3338095734343483515009979499935929'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes</a> ")
					response.Write("|<a href='http://www.co.steele.mn.us/Property_Data_Tax_Search.html' class='uLink'>Map </a>")
					response.Write("<a href='http://www.co.steele.mn.us/TREASUR/RE_Tax_Pmt_Options.html'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes</a> ")
					end if

				If cName = "Douglas" Then
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|   |")
				'	response.Write("<a href='http://morris.state.mn.us/dcmap/' class='ulink'>View Maps </a>")
				'	'response.Write("<a href='javascript:zoomto_selectfeature('strPIN');' class='uLink'> View Maps </a>")
				'   response.Write("|<a href='http://ipaddress/douglas_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> Maps </a>")
				 response.Write("|<a href='http://206.145.187.195/douglas_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> View Maps </a>")
				 Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447712482162873&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By Credit Card</a> ")
				'else
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				end if
				If cName = "Rock" Then
				 'response.Write("|<a href='http://206.145.187.195/rock_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> View Maps </a>")
				response.write("|<a href='http://rock.houstoneng.com/rock/rock.html?call=ItemQuery&buffer=500&layers=Parcels&qstring=" & strPID & "'class='ulink'> View Maps </a>")
				end if

			Case 4
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>Unpaid Tax</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>   ")

				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>|")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>|")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>|")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>|")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='hLink'>History</a>")
					If cName = "Swift" then
					Response.Write("<a href='https://swp.paymentsgateway.net/default.aspx?pg_api_login_id=X6p8j5PGd0&pg_consumerorderid=" & strPID & "&pg_billto_postal_name_first=" & firstname &"&pg_billto_postal_name_last=" & lastname &"' ) class='tLink2' align='right'>          E-payment    </a> |  ")
					'response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=222859866205626577232622909126840'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?clientId=222859866205626577232605729257656' class='tlink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     Pay By Credit Card</a> ")
					'response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=222859866205626577232622909126840&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
					end if
					If cName = "Renville" then
					'Response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Kandiyohi" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=161738092023567529211240682237982904&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Lyon" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764220341208479929&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Pope" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
					end if
					If cName = "Redwood" then
					Response.Write("<a href='http://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'> Pay By Credit Card</a> ")
					Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>  Pay By echeck</a> ")
					'Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>  Pay By echeck</a> ")
					end if
					If cName = "Yellow Medicine" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=34971905467269240275222972620077240&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Steele" then

					'response.Write("<a href='https://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=3338095734343483515010086874118329'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?clientId=3338095734343483515009979499935929'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes</a> ")
					response.Write("|<a href='http://www.co.steele.mn.us/Property_Data_Tax_Search.html' class='uLink'>Map </a>")
					response.Write("<a href='http://www.co.steele.mn.us/TREASUR/RE_Tax_Pmt_Options.html'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes</a> ")
					end if

				If cName = "Douglas" Then
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|   |")
				'	response.Write("<a href='http://morris.state.mn.us/dcmap/' class='ulink'>View Maps </a>")
				'	'response.Write("<a href='javascript:zoomto_selectfeature('strPIN');' class='uLink'> View Maps </a>")
				'   response.Write("<a href='http://ipaddress/douglas_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> Maps </a>")
				 response.Write("|<a href='http://206.145.187.195/douglas_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> View Maps </a>")
				 Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447712482162873&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By Credit Card</a> ")
				'else
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				end if
				If cName = "Rock" Then
				 'response.Write("|<a href='http://206.145.187.195/rock_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> View Maps </a>")
				response.write("|<a href='http://rock.houstoneng.com/rock/rock.html?call=ItemQuery&buffer=500&layers=Parcels&qstring=" & strPID & "'class='ulink'> View Maps </a>")
				end if

			Case 5
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Unpaid Tax</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>History</a>   ")

				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>|")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>|")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>|")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>|")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
					If cName = "Swift" then
					Response.Write("<a href='https://swp.paymentsgateway.net/default.aspx?pg_api_login_id=X6p8j5PGd0&pg_consumerorderid=" & strPID & "&pg_billto_postal_name_first=" & firstname &"&pg_billto_postal_name_last=" & lastname &"' ) class='tLink2' align='right'>          E-payment    </a> |  ")
					'response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=222859866205626577232622909126840'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?clientId=222859866205626577232605729257656' class='tlink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     Pay By Credit Card</a> ")
					'response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=222859866205626577232622909126840&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
					end if
					If cName = "Renville" then
					'Response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Kandiyohi" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=161738092023567529211240682237982904&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Lyon" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764220341208479929&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Pope" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
					end if
					If cName = "Redwood" then
					Response.Write("<a href='http://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>      Pay By Credit Card</a> ")
					Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>  Pay By echeck</a> ")
					'Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>   Pay By echeck</a> ")
					end if
					If cName = "Yellow Medicine" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=34971905467269240275222972620077240&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Steele" then

					'response.Write("<a href='https://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=3338095734343483515010086874118329'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?clientId=3338095734343483515009979499935929'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes</a> ")
					response.Write("|<a href='http://www.co.steele.mn.us/Property_Data_Tax_Search.html' class='uLink'>Map </a>")
					response.Write("<a href='http://www.co.steele.mn.us/TREASUR/RE_Tax_Pmt_Options.html'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes</a> ")
					end if

				If cName = "Douglas" Then
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|   |")
				'	response.Write("<a href='http://morris.state.mn.us/dcmap/' class='ulink'>View Maps </a>")
				'	'response.Write("<a href='javascript:zoomto_selectfeature('strPIN');' class='uLink'> View Maps </a>")
				'   response.Write("<a href='http://ipaddress/douglas_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> Maps </a>")
				 response.Write("|<a href='http://206.145.187.195/douglas_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> View Maps </a>")
				 Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447712482162873&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By Credit Card</a> ")
				'else
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				end if
				If cName = "Rock" Then
				'response.Write("|<a href='http://206.145.187.195/rock_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> View Maps </a>")
				response.write("|<a href='http://rock.houstoneng.com/rock/rock.html?call=ItemQuery&buffer=500&layers=Parcels&qstring=" & strPID & "'class='ulink'> View Maps </a>")
				end if

			Case 6
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Unpaid Tax</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>History</a>   ")

				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0' class='uLink'>General Information</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1' class='uLink'>Tax Information</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2' class='uLink'>Current Receipts</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3' class='uLink'>Unpaid Tax</a>   |   ")
				'response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>   ")
					If cName = "Swift" then
					Response.Write("<a href='https://swp.paymentsgateway.net/default.aspx?pg_api_login_id=X6p8j5PGd0&pg_consumerorderid=" & strPID & "&pg_billto_postal_name_first=" & firstname &"&pg_billto_postal_name_last=" & lastname &"' ) class='tLink2' align='right'>          E-payment    </a> |  ")
					'response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=222859866205626577232622909126840'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?clientId=222859866205626577232605729257656' class='tlink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     Pay By Credit Card</a> ")
					'response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=222859866205626577232622909126840&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     Pay By Credit Card</a> ")
					'response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
					end if
					If cName = "Renville" then
					'Response.Write("<a href='http://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=239356378198941614223105960752969913&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Kandiyohi" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=161738092023567529211240682237982904&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Lyon" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764220341208479929&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Pope" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					Response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
					end if
					If cName = "Redwood" then
					Response.Write("<a href='http://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>      Pay By Credit Card</a> ")
					Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>  Pay By echeck</a> ")
					'Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>   Pay By echeck</a> ")
					end if
					If cName = "Yellow Medicine" then
					Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=34971905467269240275222972620077240&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay By Credit Card</a> ")
					end if
					If cName = "Steele" then

					'response.Write("<a href='https://staging.officialpayments.com/pc_entry_cobrand.jsp?productId=3338095734343483515010086874118329'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes Now </a> ")
					'response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?clientId=3338095734343483515009979499935929'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes</a> ")
					response.Write("|<a href='http://www.co.steele.mn.us/Property_Data_Tax_Search.html' class='uLink'>Map </a>")
					response.Write("<a href='http://www.co.steele.mn.us/TREASUR/RE_Tax_Pmt_Options.html'  class='tLink1' align='right'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Pay Taxes</a> ")
					end if

				If cName = "Douglas" Then
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>|   |")
				'	response.Write("<a href='http://morris.state.mn.us/dcmap/' class='ulink'>View Maps </a>")
				'	'response.Write("<a href='javascript:zoomto_selectfeature('strPIN');' class='uLink'> View Maps </a>")
				'   response.Write("<a href='http://ipaddress/douglas_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> Maps </a>")
				 response.Write("|<a href='http://206.145.187.195/douglas_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> View Maps </a>")
				 Response.Write("<a href='https://www.officialpayments.com/pc_entry_cobrand.jsp?productId=4834979930617109156764447712482162873&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By Credit Card</a> ")
				'else
				'	response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4' class='uLink'>History</a>")
				end if
				If cName = "Rock" Then
				'response.Write("|<a href='http://206.145.187.195/rock_mapmorph/mapmorph/index.html?pin=" & strPID & "' class='ulink'> View Maps </a>")
				response.write("|<a href='http://rock.houstoneng.com/rock/rock.html?call=ItemQuery&buffer=500&layers=Parcels&qstring=" & strPID & "'class='ulink'> View Maps </a>")
				end if

			End Select
		%>
		</td>
	</tr>
		<tr>
			<td width="10"></td>
			<td width="580" align="right" colspan="2">
		<%
		If cName = "Pope" then
		'Response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447360294844601&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
		end if
		If cName = "Redwood" then
		'Response.Write("<a href='http://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=4834979930617109156764447463374059705&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>  Pay By echeck</a> ")
		end if
		If cName = "Swift" then
		'response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=302186245663569322297779963061598393&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
		end if
		If cName = "Yellow Medicine" then
		Response.Write("<a href='https://www.officialpayments.com/echeck/pc_entry_cobrand.jsp?productId=34971905467269240275222972620077240&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>Pay By echeck</a> ")
		end if

			%>
			</td>
	</tr>
	<tr>
		<td width="1800" bgcolor="#000000" height="1" colspan="2"></td>
	</tr>
	<tr valign="top">
		<td height="15"></td>
	</tr>
	<tr valign="top">
		<%
		if varcid = 53 or varcid = 42 or varcid = 61 or varcid = 67 or varcid = 84 or varcid = 21 then ' these are for Value Counties  only . These add a description to the page. Like PLZ or PLAT or Green Acres.....'
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
			'if didit <> 3 then	'	if objRSV3("Text3") <> ""  then 'and  objRSV3("Text2") <> ""  and objRSV3("Text3") ="" then ' and objRSV3("Text4") ="" and objRSV3("Text5") =""  then
			'Response.Write("<td width='450' align=left class='tlink2' colspan='2'> " & objRSV3("Text1") & " , " & objRSV3("Text2") & "  , " & objRSV3("Text3") & "  " & objRSV3("Text4") & "  " & objRSV3("Text5") & "</td>")
			'didit = 3
			'end if
			'end if
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
			'if didit <> 3 then
			'if objRSV2("Text1") <> "" and  objRSV2("Text2") <>"" and objRSV2("Text3") <>"" and objRSV2("Text4") <>"" and objRSV2("Text5") <>""  then
			'Response.Write("<td width='450' align=left class='tlink2' colspan='2'> " & objRSV2("Text1") & " , " & objRSV2("Text2") & " ,  " & objRSV2("Text3") & " , " & objRSV2("Text4") & " , " & objRSV2("Text5") & "</td>")
			'didit = 3
			'end if
			'end if
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
		<!-- #include file="GeneralTax2.asp" -->
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
		<!-- #include file="SpecialAssessment_Tax.asp" -->
		<%
			Case 4
		%>
		<!-- #include file="UnpaidTax.asp" -->
		<%
			Case 5
		%>
		<!-- #include file="History.asp" -->
		<%
			Case 6
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

If varcid = 42 or varcid = 21 or varcid = 53 or varcid = 61  or varcid = 84 or varcid = 67  then
	Response.Write("<a href='/tax/SearchInputReturn_value.asp?cid=" & varcid & "' class='tlink'>Another Search    |</a>&nbsp;&nbsp;&nbsp;&nbsp")
	Response.Write("<a href='/tax/ParcelListReturn_value.asp?cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='tlink'>Back to ParcelList    |</a>&nbsp;&nbsp;&nbsp;&nbsp")
	'Response.Write("<a href='/Loren/ParcelListReturn.asp?cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='tlink'>Back to ParcelList    |</a>&nbsp;&nbsp;&nbsp;&nbsp")
elseif cName = "Steele" then
	Response.Write("<a href='/website/tax/SearchInputReturn_value.asp?cid=" & varcid & "' class='tlink'>Another Search    |</a>&nbsp;&nbsp;&nbsp;&nbsp")
	Response.Write("<a href='/website/tax/ParcelListReturn_value.asp?cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='tlink'>Back to ParcelList    |</a>&nbsp;&nbsp;&nbsp;&nbsp")

Else
	Response.Write("<a href='/tax/searchinputreturn.asp?cid=" & varcid & "' class='tlink'>Another Search    |</a>&nbsp;&nbsp;&nbsp;&nbsp")
	'Response.Write("<a href='/Loren/searchinputreturn.asp?cid=" & varcid & "' class='tlink'>Another Search    |</a>&nbsp;&nbsp;&nbsp;&nbsp")
	Response.Write("<a href='/tax/ParcelListReturn.asp?cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='tlink'>Back to ParcelList    |</a>&nbsp;&nbsp;&nbsp;&nbsp")
	'Response.Write("<a href='/Loren/ParcelListReturn.asp?cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='tlink'>Back to ParcelList    |</a>&nbsp;&nbsp;&nbsp;&nbsp")
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
<%
			'If cName = "Steele" Then
			'	Response.Write("<a href='http://www.co.steele.mn.us/' class='tlink'>Steele County Home Page</a>")
			'End If

%>
			</td>

	<tr>
		<td height="15" colspan="2"></td>
	</tr>
	<tr valign="top">
		<td width="10"></td>
		<td width="650" align="right" class="STitle2">

			</td>


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




objRS.Close
Set objRS = Nothing

'objRScount.Close
'Set objRScount = Nothing

%>
<%
'Response.Write("Your session county at : " & cid )

%>



</body>
</html>
