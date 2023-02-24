<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Parcel Search Results</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
	Dim cid, cName, sesTotal
	cid = request.QueryString("cid")
	intParcelNo = request.QueryString("varintParcelNo")
	strAddress = request.QueryString("varstrAddress")
	strName = request.QueryString("varstrName")
	intSect = request.QueryString("varintSect")
	intTwp = request.QueryString("varintTwp")
	intRange = request.QueryString("varintRange")
   // strAmount=request.QueryString("payAmount")

	response.AddHeader "Access-Control-Allow-Origin", "*"
	sesTotal = Session("amtTotal")
	response.Write("<link rel='stylesheet' href='" & cid & ".css' type='text/css'>")

		Select Case cid
			Case 13
			cName = "Dunn"
			Case 23
			cName = "LaMoure"
			Case 27
			cName = "McKenzie"
			Case 30
			cName = "Morton"
			Case 31
			cName = "Mountrail"
			Case 34
			cName = "Pembina"
			Case 02
			cName = "Barnes"
			Case 37
			cName = "Ransom"
			Case 41
			cName = "Sargent"
			Case 47
			cName = "Stutsman"
		End Select


%>
<script type="text/javascript">

//tmpcid = <%=Session("CountyID")%>;

//strPID = <%=request.QueryString("pid")%>;
function createCORSRequest(method, url) {
  var xhr = new XMLHttpRequest();
  if ("withCredentials" in xhr) {

    // Check if the XMLHttpRequest object has a "withCredentials" property.
    // "withCredentials" only exists on XMLHTTPRequest2 objects.
    xhr.open(method, url, true);

  } else if (typeof XDomainRequest != "undefined") {

    // Otherwise, check if XDomainRequest.
    // XDomainRequest only exists in IE, and is IE's way of making CORS requests.
    xhr = new XDomainRequest();
    xhr.open(method, url);

  } else {

    // Otherwise, CORS is not supported by the browser.
    xhr = null;

  }
  return xhr;
}
function createRequestObject() 
{
        if (window.XMLHttpRequest) 
        {
                return xmlhttprequest = new XMLHttpRequest(); 
        } 
      else if (window.ActiveXObject) 
      {  
            return xmlhttprequest = new ActiveXObject("Microsoft.XMLHTTP"); 
      } 
}

</script>
</head>

<!-- #include file="insDB.asp" -->

<body>
<%
Function printTaxRecord()
	Dim recordNumber
	For recordNumber=1 to 5
	Session("amtTotal") = 0
	sesTotal = 0
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
				If (objRS3(("RCAMT1" & recordNumber )) <> 0) Then
					Response.Write("<td width='60' class='rText' align='right'>" & objRS3(("RCTYP"  & "1" & recordNumber)) & "</td>")
					Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCAMT"  & "1" & recordNumber)), 2) & "</td>")
				Else
					Response.Write("<td width='60' class='rText'>&nbsp;</td>")
					Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				End If
				sesTotal = objRS3("RCAMT" & "1" & recordNumber)
				If (objRS3("RCSAA1" & recordNumber) <> 0) Then
					Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAR"  & "1" & recordNumber)) & "</td>")
					Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAC" & "1" & recordNumber)) & "</td>")
					Response.Write("<td width='80' class='rText' align='right'>" & objRS3(("RCSAA" & "1" & recordNumber)) & "</td>")
				else
					Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
					Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
					Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
				end if


				If objRS3(("RCTYP"  & "1" & recordNumber)) = "DISC" Then
				else
				sesTotal = sesTotal + (objRS3("RCSAA" & "1" & recordNumber))
				end if

				Session("recnumberend") = recordNumber
				Response.Write("</tr>")
			printTaxValues(recordNumber)
		Session("amtTotal") = 0
		sesTotal = 0
		End If
	Next
End Function
Function calcZeroTnocomma(strData, intPlaces)
	If objRS(strData) <> "" Then
		strData = (FormatNumber(objRS(strData), intPlaces,,,0))
	Else
		strData = "0.00"
	End If
	calcZeroTnocomma = strData
end Function

Function printTaxValues(recordNumber)
	Dim subrecordNumber, tempsubrecordnumber
	For subrecordNumber=2 to 6
		If (objRS3(("RCSAA" & subrecordNumber & recordNumber )) <> 0) or (objRS3(("RCAMT" & subrecordNumber & recordNumber))<>0) Then

			If objRS3(("RCTYP"  & subrecordNumber & recordNumber)) = "DISC" Then
			else
				sesTotal = sesTotal + objRS3("RCAMT" & subrecordNumber & recordNumber )
			end if

			Response.Write("<tr valign='top'>")
			If (objRS3(("RCAMT" & subrecordNumber & recordNumber )) <> 0) Then
				Response.Write("<td width='60' class='rText' align='right'>"  & objRS3(("RCTYP" & subrecordNumber & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS3(("RCAMT" & subrecordNumber & recordNumber)), 2) & "</td>")
			Else
				Response.Write("<td width='60' class='rText'>&nbsp;</td>")
				Response.Write("<td width='80' class='rText' align='right'>&nbsp;</td>")
			End If
			If (objRS3("RCSAA" & subrecordNumber & recordNumber) <> 0) Then
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
			Exit For	
		End If
	Next
	If (subrecordNumber > 4) Then
		Response.Write("<td width='260' class='rText' align='left' rowspan='6'>&nbsp;</td>")
	End If
	Response.Write("<td width='60' class='rText'>&nbsp;</td>")
	Response.Write("<td width='80' class='rText'>&nbsp;</td>")
	Response.Write("<td width='80' class='rText'>&nbsp;</td>")
	Response.Write("<td width='80' class='rText'>Total</td>")
	Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber((sesTotal), 2) & "</td>")

	End Function

Function printTaxRecord4()
	Dim recordNumber
	For recordNumber=6 to 10
		If objRS4("RCBTDT" & recordNumber) = 0 Then
			Response.Write("<tr valign='top'>")
				If recordNumber = 1 Then
					Response.Write("<td class='rText' align='left' rowspan='7'>No Tax Receipt Information</td>")
				End If
			recordNumber = 10
		Else
			Response.Write("<tr valign='top'>")
				Response.Write("<td width='260' class='rText' align='left' rowspan='6'>")
				Response.Write("<b>" & calcDateRS4(("RCBTDT"), recordnumber) & "</b><br>")
				Response.Write("<b>Batch # " & objRS4(("RCBAT#" & recordNumber)) & "</b><br>")
				Response.Write("<b>Paid</b> by " & objRS4(("RCPDBY" & recordNumber)) & "<br>")
				Response.Write("<b>Validation #</b> " & objRS4(("RCVAL#" & recordNumber)))
				Response.Write("</td>")
				Response.Write("<td width='60' class='rText' align='right'>" & objRS4(("RCTYP"  & "1" & recordNumber)) & "</td>")

				If objRS4(("RCTYP"  & "1" & recordNumber)) = "DISC" Then
				else
				sesTotal = sesTotal + (objRS4("RCSAA" & "1" & recordNumber))
				end if


				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS4(("RCAMT"  & "1" & recordNumber)), 2) & "</td>")
				sesTotal = objRS4("RCAMT" & "1" & recordNumber)
				Response.Write("<td width='80' class='rText' align='right'>" & objRS4(("RCSAR"  & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & objRS4(("RCSAC" & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & objRS4(("RCSAA" & "1" & recordNumber)) & "</td>")
				Session("recnumberend") = recordNumber
				Response.Write("</tr>")
			printTaxValues4(recordNumber)
		End If
	Next
End Function

'*******   Function added here to take the Table 4 records print them to the screen.
Function printTaxValues4(recordNumber)
	Dim subrecordNumber, tempsubrecordnumber
	For subrecordNumber=2 to 6
		If (objRS4(("RCSAA" & subrecordNumber & recordNumber )) > 0) or (objRS4(("RCAMT" & subrecordNumber & recordNumber))>0) Then
				Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText' align='right'>"  & objRS4(("RCTYP" & subrecordNumber & recordNumber)) & "</td>")

				If objRS4(("RCTYP"  & subrecordNumber & recordNumber)) = "DISC" Then
				else
				sesTotal = sesTotal + objRS4("RCAMT" & subrecordNumber & recordNumber )
				end if



				If objRS4(("RCAMT" & subrecordNumber & recordNumber )) > 0 Then
					Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS4(("RCAMT" & subrecordNumber & recordNumber)), 2) & "</td>")
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

				If objRS6(("RCTYP"  & "1" & recordNumber)) = "DISC" Then
				else
				sesTotal = sesTotal + (objRS6("RCSAA" & "1" & recordNumber))
				end if

				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS6(("RCAMT"  & "1" & recordNumber)), 2) & "</td>")
				sesTotal = objRS6("RCAMT" & "1" & recordNumber)
				Response.Write("<td width='80' class='rText' align='right'>" & objRS6(("RCSAR"  & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & objRS6(("RCSAC" & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS6(("RCSAA" & "1" & recordNumber)), 2) & "</td>")
				Session("recnumberend") = recordNumber
				Response.Write("</tr>")
			printTaxValues6(recordNumber)
		End If
	Next
End Function


Function printTaxValues6(recordNumber)
	Dim subrecordNumber, tempsubrecordnumber
	For subrecordNumber=2 to 6
		If (objRS6(("RCSAA" & subrecordNumber & recordNumber )) > 0) or (objRS6(("RCAMT" & subrecordNumber & recordNumber))>0) Then
				Response.Write("<tr valign='top'>")
				Response.Write("<td width='60' class='rText' align='right'>"  & objRS6(("RCTYP" & subrecordNumber & recordNumber)) & "</td>")

				If objRS6(("RCTYP"  & subrecordNumber & recordNumber)) = "DISC" Then
				else
				sesTotal = sesTotal + objRS6("RCAMT" & subrecordNumber & recordNumber )
				end if


				If objRS6(("RCAMT" & subrecordNumber & recordNumber )) > 0 Then
					Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS6(("RCAMT" & subrecordNumber & recordNumber)), 2) & "</td>")
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

				If objRS7(("RCTYP"  & "1" & recordNumber)) = "DISC" Then
				else
				sesTotal = sesTotal + (objRS4("RCSAA" & "1" & recordNumber))
				end if


				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS7(("RCAMT"  & "1" & recordNumber)), 2) & "</td>")
				sesTotal = objRS7("RCAMT" & "1" & recordNumber)
				Response.Write("<td width='80' class='rText' align='right'>" & objRS7(("RCSAR"  & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & objRS7(("RCSAC" & "1" & recordNumber)) & "</td>")
				Response.Write("<td width='80' class='rText' align='right'>" & FormatNumber(objRS7(("RCSAA" & "1" & recordNumber)), 2) & "</td>")
				Session("recnumberend") = recordNumber
				Response.Write("</tr>")
			printTaxValues7(recordNumber)
		End If
	Next
End Function


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

Function calcZeroRS5(strData, intPlaces)
	If objRS5(strData) <> "" Then
		strData = (FormatNumber(objRS5(strData), intPlaces))
	Else
		strData = "0.00"
	End If
	calcZeroRS5 = strData
end Function

Function calcZeroRS52(strData, intPlaces)
	If objRS52(strData) <> "" Then
		strData = (FormatNumber(objRS52(strData), intPlaces))
	Else
		strData = "0.00"
	End If
	calcZeroRS52 = strData
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

Function calcDateRS5(strData)
	Dim strYear, strMonth, strDay, intLength
	strData = objRS5("UPINDT")
	intLength = Len(strData)
	strYear = Right(strData, 4)
	If intLength = 7 Then
		strMonth = Left(strData, 1)
		strDay = Mid(strData, 2, 2)
	Else
		strMonth = Left(strData, 2)
		strDay = Mid(strData, 3, 2)
	End If
	calcDateRS5 = strMonth + "/" + strDay + "/" + strYear
end Function

Function calcDateMoYrRS5(strData)
	Dim strYear, strMonth, strDay, intLength
	strData = objRS5("UPINDT")
	intLength = Len(strData)
	strYear = Right(strData, 4)
	If intLength = 7 Then
		strMonth = Left(strData, 1)
		strDay = Mid(strData, 2, 2)
	Else
		strMonth = Left(strData, 2)
		strDay = Mid(strData, 3, 2)
	End If
	calcDateMoYrRS5 = strMonth + "/" + strYear
end Function

Function calcDateRS52(strData)
	Dim strYear, strMonth, strDay, intLength
	strData = objRS52("UPINDT")
	intLength = Len(strData)
	strYear = Right(strData, 4)
	If intLength = 7 Then
		strMonth = Left(strData, 1)
		strDay = Mid(strData, 2, 2)
	Else
		strMonth = Left(strData, 2)
		strDay = Mid(strData, 3, 2)
	End If
	calcDateRS52 = strMonth + "/" + strDay + "/" + strYear
end Function

Function calcDateMoYrRS52(strData)
	Dim strYear, strMonth, strDay, intLength
	strData = objRS52("UPINDT")
	intLength = Len(strData)
	strYear = Right(strData, 4)
	If intLength = 7 Then
		strMonth = Left(strData, 1)
		strDay = Mid(strData, 2, 2)
	Else
		strMonth = Left(strData, 2)
		strDay = Mid(strData, 3, 2)
	End If
	calcDateMoYrRS52 = strMonth + "/" + strYear
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


' ****   the calcTotRS5   function goes here.   * * * * *
Function calcTotRS5(strData)
	If objRS5("UPTOTDUE5") > 0 Then
	strData = objRS5("UPTOTDUE1")  + objRS5("UPTOTDUE2") + objRS5("UPTOTDUE3") + objRS5("UPTOTDUE4") + objRS5("UPTOTDUE5")
	elseif 	objRS5("UPTOTDUE4") > 0 Then
	strData = objRS5("UPTOTDUE1")   + objRS5("UPTOTDUE2") + objRS5("UPTOTDUE3") + objRS5("UPTOTDUE4")
	elseif   objRS5("UPTOTDUE3") > 0 Then
	strData = objRS5("UPTOTDUE1")   + objRS5("UPTOTDUE2") + objRS5("UPTOTDUE3")
	elseif objRS5("UPTOTDUE2") > 0 Then
	strData = objRS5("UPTOTDUE1")   + objRS5("UPTOTDUE2")
	elseif objRS5("UPTOTDUE1") > 0 Then
 	strData = objRS5("UPTOTDUE1")
	else
	strData = "0.00"
	end if
	calcTotRS5 = strData
end Function

Function calcTotRS52(strData2)
	If objRS52("UPTOTDUE5") > 0 Then
	strData2 = objRS52("UPTOTDUE1")  + objRS52("UPTOTDUE2") + objRS52("UPTOTDUE3") + objRS52("UPTOTDUE4") + objRS52("UPTOTDUE5")
	elseif 	objRS52("UPTOTDUE4") > 0 Then
	strData2 = objRS52("UPTOTDUE1")   + objRS52("UPTOTDUE2") + objRS52("UPTOTDUE3") + objRS52("UPTOTDUE4")
	elseif   objRS52("UPTOTDUE3") > 0 Then
	strData2 = objRS52("UPTOTDUE1")   + objRS52("UPTOTDUE2") + objRS52("UPTOTDUE3")
	elseif objRS52("UPTOTDUE2") > 0 Then
	strData2 = objRS52("UPTOTDUE1")   + objRS52("UPTOTDUE2")
	elseif objRS52("UPTOTDUE1") > 0 Then
 	strData2 = objRS52("UPTOTDUE1")
	else
	strData2 = "0.00"
	end if
	calcTotRS52 = strData2
end Function

Function calcZip(strData)
	If objRS(strData) = "00000" Then
		strData = ""
	End If
	calcZip = strData
end Function

'****  Check to see what the most current year of data is in STMT table
Set objCommand = Nothing
Set objCommand = Server.CreateObject("ADODB.Command")
objCommand.ActiveConnection = strConnect
objCommand.CommandText="SELECT RIGHT(name,2) as tempyr FROM MsysObjects where type=1 and LEFT(name,6)='TXSTMT' order by RIGHT(name,2) DESC "

objCommand.CommandType=1
Dim current,previous

'Set objRSTYR = objCommand.Execute
'current=objRSTYR("tempyr")

	' objRSTYR.MoveNext
'previous=objRSTYR("tempyr")

'Response.Write('year='& current)
'Response.Write('year2=' & previous)



Dim objCommand, objRS, strQueryString, strPID, strTID, objRS3, objRS5, objRS52, objRScount, intnumberQ

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strQueryString & "'"

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

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 3 - Part1 of RCPT Sets(1-5)] " & strQueryString
objCommand.CommandType = 1

Set objRS3 = objCommand.Execute
Set objCommand = Nothing

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
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
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 5 - Unpaid Taxes].TXPRCL = '" & strQueryString & "'"

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 5 - Unpaid Taxes] " & strQueryString
objCommand.CommandType = 1

Set objRS5 = objCommand.Execute
Set objCommand = Nothing

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE ((([Table 5 - Unpaid Taxes].TXPRCL) = '" & strQueryString & "') AND (([Table 5 - Unpaid Taxes].TXREC)=2))"

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 5 - Unpaid Taxes] " & strQueryString
objCommand.CommandType = 1

Set objRS52 = objCommand.Execute
Set objCommand = Nothing



Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE [Table 6 - Part3 of RCPT Sets(11-15)].TXPRCL = '" & strQueryString & "'"

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 6 - Part3 of RCPT Sets(11-15)] " & strQueryString
objCommand.CommandType = 1

Set objRS6 = objCommand.Execute
Set objCommand = Nothing

Set objCommand = Server.CreateObject("ADODB.Command")
strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
intYYR = Session("intYEAR")
strQueryString = request.QueryString("pid")

strQueryString = "WHERE (([Table 2 - Special/Ditch Info].TXPRCL) = '" & strQueryString & "')"

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 2 - Special/Ditch Info] " & strQueryString
objCommand.CommandType = 1
Set objRSSPT = objCommand.Execute
Set objCommand = Nothing


Set objCommand = Server.CreateObject("ADODB.Command")

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 8 - Misc Data];"
objCommand.CommandType = 1

Set objRSCNTDT = objCommand.Execute

Set objCommand = Nothing


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

	Dim sesPID
	Session("sesPID") = concstrPID
	cesPID = Session("sesPID")



'CREATE THE OBJRSV2 FROM TABLE 9 **********************
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
' End of Create the OBJRSV2 from Table 9 * * * * * * * * * * * * * *

%>

<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="120" >As of : <b><%= objRSCNTDT("AsOfDate") %></b></td><td width="535" align="right" colspan="2">Parcel Number: <b><%= objRS("TXPRCL") %></b></td>
	</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">

		<td  width="580"  align="right" valign="top">Payable Year: <b></b></td>
		<td width="75" align="right" class="tText3" ><%= objRS("TXYEAR") %></td>
		<td width="3"></td>
<%
			Select Case strTID
			Case 0
'Response.Write(cid=02)

	If  cid=37 Then
		
	If objRS("TXREMH")="MOBILE HOME"  Then
					response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				Else
					Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2021&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				End If

	'else

	'If objRS("TXREMH")="MOBILE HOME"  Then
					'response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2017&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'Else
					'Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2016&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'End If

	End If
        'Barnes
        If cid=02 Then
		
	     If objRS("TXREMH")="MOBILE HOME"  Then
	          response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
	     Else
		Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2021&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
	     End If

	End If
  'Pembina
        If cid=34 Then
		
	     If objRS("TXREMH")="MOBILE HOME"  Then
	          response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
	     Else
		Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2021&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
	     End If

	End If
     'McKenzie
         If cid=27 Then
		
	     If objRS("TXREMH")="MOBILE HOME"  Then
	          response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2021&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
	     Else
		Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2021&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
	     End If

	End If
       'Dunn
        If cid=13 Then
		
	     If objRS("TXREMH")="MOBILE HOME"  Then
	          response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
	     Else
		Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2021&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
	     End If

	End If
        'Mountrail
        If cid=31 Then
		
	     If objRS("TXREMH")="MOBILE HOME"  Then
	          response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
	     Else
		Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2021&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
	     End If

	End If
 'Sargent
        If cid=41 Then
		
	     If objRS("TXREMH")="MOBILE HOME"  Then
	          response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2021&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
	     Else
		Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2020&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
	     End If

	End If


        'LaMoure
        If cid=23 Then
		
	If objRS("TXREMH")="MOBILE HOME"  Then
					response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				Else
					Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2021&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				End If


	End If


	'End If
       




	'If objRS("TXREMH")="MOBILE HOME"  Then
					'response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2017&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'Else
					'Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2016&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'End If

	
		'Select Case cid
				'Case 13
				'If objRS("TXREMH")="MOBILE HOME"  Then
					'If objRS("TXYEAR")=2014  Then
					'response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND13MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
					'End If

				'else
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND13STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'End If
				'Case 23
				'If objRS("TXREMH")="MOBILE HOME"  Then
				'	If objRS("TXYEAR")=2014  Then
					'response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND23MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
					'End If
				'Else
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND23STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'End If
				'Case 27
				'If objRS("TXREMH")="MOBILE HOME"  Then
				'	If objRS("TXYEAR")=2014  Then
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND27MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'	End If
				'Else
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND27STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'End If
				'Case 30
				'If objRS("TXREMH")="MOBILE HOME"  Then
				'	If objRS("TXYEAR")=2014  Then
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND30MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'	End If
				'Else
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND30STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'End If
				'Case 31
				'If objRS("TXREMH")="MOBILE HOME"  Then
				'	If objRS("TXYEAR")=2014  Then
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND31MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'	End If
				'Else
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND31STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'End If
				'Case 34
				'If objRS("TXREMH")="MOBILE HOME"  Then
				'	If objRS("TXYEAR")=2014  Then
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND34MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'	End If
				'Else
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND34STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'End If
				'Case 02
				'If objRS("TXREMH")="MOBILE HOME"  Then
				'	If objRS("TXYEAR")=2014  Then
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND02MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'	End If
				'Else
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND02STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'End If
				'Case 37
				'If objRS("TXREMH")="MOBILE HOME"  Then
					'If objRS("TXYEAR")=2014  Then
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND37MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'	End If
				'Else
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND37STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'End If
				'Case 41
				'If objRS("TXREMH")="MOBILE HOME"  Then
				'	If objRS("TXYEAR")=2014  Then
				'	response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND41MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'	End If
				'Else
				'	Response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND41STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'End If
				'Case 47
				'If objRS("TXREMH")="MOBILE HOME"  Then
					'If objRS("TXYEAR")=2014  Then
					'response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND47MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
					'End If
				'else
					'response.Write("<td width='200' align='right' colspan='2'><A HREF='http://ndpropertytax.org:41080/NDiText/ND47STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtPrev.gif'></A></td>")
				'End If


		'End Select
	end Select

%>
	</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="680" align="left" colspan="6">
<%

			Dim strBegin, strEnd, strMid, newPid
			strbegin=Left(strPID,2)
			strMid=Mid(strPID,4,2)
			strEnd=Right(strPID,5)
			newPid=strbegin+strMid+strEnd

			Select Case strTID
			Case 0
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Unpaid Tax</a>   |    ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>   ")
				If cName = "Barnes" then
					If objRS("TXREMH")="MOBILE HOME"  Then
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022483528542320825&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					else
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022477614372354233&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					end if
				end if
				If cName = "Morton" Then
				response.write("| <a href='http://mortonnd.mygisonline.com/?pin=" & newPid & "'class='ulink' target=_new, toolbar=no, menubar=no >View Maps </a>")
				end if
				response.write("<td width='25'>&nbsp;</td>")

				If cid=37  Then
					If objRS("TXREMH")="MOBILE HOME"  Then
						response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2023&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
					Else
						Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				End If
				'Else
				'If objRS("TXREMH")="MOBILE HOME"  Then
					'response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2018&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				'Else
					'Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2017&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				'End If
				End If
 			         If cid=02  Then
					If objRS("TXREMH")="MOBILE HOME"  Then
						response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2023&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
					Else
						Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				        End If
				
				End If


                                If cid=13  Then
					If objRS("TXREMH")="MOBILE HOME"  Then
						response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2023&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
					Else
						Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				        End If
				
				End If

                                If cid=27  Then
					If objRS("TXREMH")="MOBILE HOME"  Then
						response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
					Else
						Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				        End If
				
				End If
   				If cid=34  Then
					If objRS("TXREMH")="MOBILE HOME"  Then
						response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2023&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
					Else
						Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				        End If
				
				End If
                                If cid=31  Then
					If objRS("TXREMH")="MOBILE HOME"  Then
						response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2023&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
					Else
						Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				        End If
				
				End If

 				If cid=41  Then
					If objRS("TXREMH")="MOBILE HOME"  Then
						response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
					Else
						Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2021&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				        End If
				
				End If

                                If cid=23  Then
					If objRS("TXREMH")="MOBILE HOME"  Then
						response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2023&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
					Else
						Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2022&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				End If
				'Else
				'If objRS("TXREMH")="MOBILE HOME"  Then
					'response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2018&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				'Else
					'Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2017&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				'End If
				End If

				'If objRS("TXREMH")="MOBILE HOME"  Then
					'response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2018&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				'Else
					'Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2017&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
				'End If
                    

		'Select Case cid
		'		Case 13
		'		If objRS("TXREMH")="MOBILE HOME"  Then
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND13MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		else
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND13STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		End If
		'		Case 23
		'		If objRS("TXREMH")="MOBILE HOME"  Then
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND23MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		Else
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND23STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		End If
		'		Case 27
		'		If objRS("TXREMH")="MOBILE HOME"  Then
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND27MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		Else
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND27STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		End If
		'		Case 30
		'		If objRS("TXREMH")="MOBILE HOME"  Then
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2015&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		Else
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2014&mhid=RE' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		End If
		'		Case 31
		'		If objRS("TXREMH")="MOBILE HOME"  Then
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND31MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		Else
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND31STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		End If
		'		Case 34
		'		If objRS("TXREMH")="MOBILE HOME"  Then
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND34MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		Else
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND34STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		End If
		'		Case 02
		'		If objRS("TXREMH")="MOBILE HOME"  Then
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND02MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		Else
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND02STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		End If
		'		Case 37
		'		If objRS("TXREMH")="MOBILE HOME"  Then
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2015&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		Else
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2014&mhid=RE' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		End If
		'		Case 41
		'		If objRS("TXREMH")="MOBILE HOME"  Then
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2015&mhid=MH' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		Else
		'			Response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/NDSTMTC.jsp?cid=" & cid & "&pid=" & strPID & "&yrid=2014&mhid=RE'' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		End If
		'		Case 47
		'		If objRS("TXREMH")="MOBILE HOME"  Then
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND47MHTSC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'			Else
		'			response.Write("<td><A HREF='http://ndpropertytax.org:41080/NDiText/ND47STMTC.jsp?cid=" & cid & "&pid=" & strPID & "' target=_new, toolbar=no, menubar=no ><IMG SRC='images/PrntTxStmtCurnt.gif'></A></td>")
		'		End If
		'End Select


			Case 1
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Unpaid Tax</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>   ")
				If cName = "Barnes" then
					If objRS("TXREMH")="MOBILE HOME"  Then
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022483528542320825&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					else
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022477614372354233&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					end if
				end if
				If cName = "Morton" Then
					response.write("|  <a href=http://mortonnd.mygisonline.com/?pin=" & newPid & "'class='ulink'>View Maps </a>")
				end if
				If cName = "Mountrail" then
					Dim payAmount
					payAmount=FormatNumber(objRS5("UPTOTDUE1"), 2)
					Session("payAmount") = payAmount
					Response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=7&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "&payAmount=" & payAmount & "' class='tLink2'>Electronic Payment</a>   ")					
				end if
			Case 2
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Unpaid Tax</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>   ")
                               If cName = "Barnes" then
					If objRS("TXREMH")="MOBILE HOME"  Then
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022483528542320825&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					else
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022477614372354233&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					end if
				end if
				If cName = "Morton" Then
					response.write("|  <a href='http://mortonnd.mygisonline.com/?pin=" & newPid & "'class='ulink'>View Maps </a>")
				end if
			Case 3
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Unpaid Tax</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>   ")
If cName = "Barnes" then
					If objRS("TXREMH")="MOBILE HOME"  Then
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022483528542320825&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					else
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022477614372354233&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					end if
				end if
				If cName = "Morton" Then
					response.write("|  <a href='http://mortonnd.mygisonline.com/?pin=" & newPid & "'class='ulink'>View Maps </a>")
				end if
			Case 4
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>Unpaid Tax</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>History</a>   ")
If cName = "Barnes" then
					If objRS("TXREMH")="MOBILE HOME"  Then
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022483528542320825&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					else
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022477614372354233&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					end if
				end if
				If cName = "Morton" Then
					response.write("|  <a href='http://mortonnd.mygisonline.com/?pin=" & newPid & "'class='ulink'>View Maps </a>")
				end if
If cName = "Mountrail" then
					
					payAmount=FormatNumber(objRS5("UPTOTDUE1"), 2)
					Session("payAmount") = payAmount
					Response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=7&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "payAmount=" & payAmount & "' class='tLink2'>Electronic Payment</a>   ")					
				end if
			Case 5
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Unpaid Tax</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>History</a>   ")
If cName = "Barnes" then
					If objRS("TXREMH")="MOBILE HOME"  Then
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022483528542320825&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					else
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022477614372354233&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					end if
				end if
				If cName = "Morton" Then
					response.write("|  <a href='http://mortonnd.mygisonline.com/?pin=" & newPid & "'class='ulink'>View Maps </a>")
				end if

			Case 6
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>General Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Tax Info</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=2&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Current Receipts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=3&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Special Asmts</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=4&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='uLink'>Unpaid Tax</a>   |   ")
				response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=5&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='hLink'>History</a>   ")
If cName = "Barnes" then
					If objRS("TXREMH")="MOBILE HOME"  Then
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022483528542320825&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					else
						Response.Write("<a href='https://www.officialpayments.com/pc_entry_standard.jsp?productId=601186845427044431022477614372354233&cde-ParcNumb-0=" & strPID & "'  class='tLink2' align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
					end if
				end if
				If cName = "Morton" Then
					response.write("|  <a href='http://mortonnd.mygisonline.com/?pin=" & newPid & "'class='ulink'>View Maps </a>")
				end if
			End Select
		%>
		</td>
	</tr>
		<tr>
			<td width="10"></td>
			<td width="580" align="right" colspan="2">
			</td>
	</tr>
	<tr>
		<td width="650" bgcolor="#000000" height="1" colspan="2"></td>
	</tr>
	<tr valign="top">
		<td height="15"></td>
	</tr>
	<tr valign="top"></tr>

	<tr valign="top">
		<td width="10"></td>
		<td width="680">
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
			Case 7
		%>
		<!-- #include file="Payment.asp" -->
		<%
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
	<td height="15" colspan="3"></td>
	</tr>
	<td width="1175"  align="Right" class="STitle2">
<%
			Response.Write("<a href='searchinput.asp?cid=" & cid & "' class='tlink'>Another Search    |</a>&nbsp;&nbsp;")
			Response.Write("<a href='ParcelListReturn.asp?pid=" & strPID & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "' class='tlink'>Back to ParcelList    |</a>&nbsp;&nbsp")
%>

	</td>
	<tr>
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
		Lamoure County<br>
		PO Box 128<br>
		Lamoure ND 58458-0128<br>
			</td>
		<td width="200"  class="STitle1">
		Treasurer<br>
		Lamoure County<br>
		PO Box 122<br>
		Lamoure ND 58458-0122<br>
		701-883-6090
			</td>
<%
	end if
%>
<%
	if cid = 37 then
%>
		<td width="180"  class="STitle1">
		Auditor<br>
		Ransom County Treasurer<br>
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
		Pam Maloney<br>
		355 Main St S, Suite 1<br>
		Forman ND 58032-4149<br>
			</td>
		<td width="200"  class="STitle1">
		Treasurer<br>
		Alison Toepke<br>
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