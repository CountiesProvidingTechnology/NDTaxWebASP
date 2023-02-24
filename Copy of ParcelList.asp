</SCRIPT>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Parcel Search Results</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
	Dim varcid
	'cid = Session("CountyID")
	'response.Write(" the valuePL1 of varcid ; " & varcid )
	varcid = Request.Form("cid")
	'response.Write(" the valuePL2 of varcid ; " & varcid )
	response.Write("<link rel='stylesheet' href='" & varcid & ".css' type='text/css'>")

	'response.Write (" Session Var " & varcid  )
%>
</head>
<!-- #include file="insDB.asp" -->
<body>
<%
If cName = "Douglas" Then
	Response.Write("<input name='returnHomeButton' type='button' value='Return to Douglas Home Page' onClick=window.location='http://www.co.douglas.mn.us'>")
End If
Dim objCommand, objRS, intParcelNo, strAddress, Addr, strQueryString, intSet, strTitle1, strTitle2, strTitle3, strTitle4, strName, intPrclNo, strAddr, strNam
intParcelNo = Request.Form("ParcelNo")
'response.Write(" the value of intParcelNo ; " & intParcelNo )
'Session("PrclNo") = intParcelNo
strAddress = Request.Form("pAddress")
strSAddress = Request.Form("sAddress")
'Session("Addr") = strAddress
'Session("sAddr") = strSAddress
Addr = Session("Addr")
strName = Request.Form("fName")
'Session("Nam") = strName

intSect = Request.Form("sName")
'Session("SecNam") = intSect
intTwp = Request.Form("tName")
'Session("TwnNam") = intTwp
intRange = Request.Form("rName")
'Session("RngNam") = intRange

'***  IF user enters a * in the search box the user is taken back to the disclaimer page. added 9-23-08 LEM
if strAddress = "*" or strSAddress = "*" then
    Response.Redirect("SearchInput.asp?cid=" & varcid & "")
End If
If strName = "*" Then
    Response.Redirect("SearchInput.asp?cid=" & varcid & "")
End If

'**  Function to format the entry of the parcel number search from the user
'**  this will allow the user to enter numbers only in the text field of the form.
'**  this is in ND web tax' Added on 9-4-08 LEM
Function fmtparcelnum(strData)
	Dim strtrimdata
	strData = Request.Form("ParcelNo")
	strtrimdata = Trim(strData)
	intLength = Len(strtrimdata)

If varcid = 21 or varcid=26 or varcid=41 or varcid=45 or varcid=53 or varcid=61 or varcid=67 or varcid=75 or varcid=76  Then         'Creates  formatted parcel XX-XXXX-XXX  (2-4-3)
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
			fmtparcelnum = strleftChars + "-" + strmidChars + "*"
	'	Response.write(" the value of the fmtparcelnum in len 4 : " & fmtparcelnum)
		elseif intlength = 5 Then
			strleftChars = Left(strtrimdata, 2)
		'	strmidChars = Mid(strtrimdata, 3, 2)
			strrightChars = Right(strtrimdata, 3)
		'	fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
			fmtparcelnum = strleftChars + "-" + strrightChars
	'	Response.write(" the value of the fmtparcelnum in len 5 : " & fmtparcelnum)
		elseif intlength = 6 Then
			strleftChars = Left(strtrimdata, 2)
		'	strmidChars = Mid(strtrimdata, 3, 2)
			strrightChars = Right(strtrimdata, 4)
		'	fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
			fmtparcelnum = strleftChars + "-" + strrightChars
		'Response.write(" the value of the fmtparcelnum in len 6 : " & fmtparcelnum)'
		elseif intLength = 7 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 4)
			strrightChars = Right(strtrimdata, 1)
			fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars + "*"
		elseif intLength = 8 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 4)
			strrightChars = Right(strtrimdata, 2)
			fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars + "*"
		elseif intlength = 9 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 4)
			strrightChars = Right(strtrimdata, 3)
			fmtparcelnum = strleftChars + "-" + strmidChars +  "-" + strrightChars
		elseif intlength = 10 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 4)
			strrightChars = Right(strtrimdata, 3)
			fmtparcelnum = strleftChars + "-" + strmidChars +  "-" + strrightChars

		End If
	End If

If varcid = 34 or varcid=48 or varcid=51 or varcid=64 or varcid=84 or varcid=87  Then         'Creates  formatted parcel XX-XXX-XXXX  (2-3-4)
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
			fmtparcelnum = strleftChars + "-" + strmidChars + "*"

		elseif intlength = 5 Then
			strleftChars = Left(strtrimdata, 2)

			strrightChars = Right(strtrimdata, 3)

			fmtparcelnum = strleftChars + "-" + strrightChars
	'	Response.write(" the value of the fmtparcelnum in len 5 : " & fmtparcelnum)
		elseif intlength = 6 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 3)
			strrightChars = Right(strtrimdata, 1)
			fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
		'	fmtparcelnum = strleftChars + "-" + strrightChars
		'Response.write(" the value of the fmtparcelnum in len 6 : " & fmtparcelnum)'
		elseif intLength = 7 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 3)
			strrightChars = Right(strtrimdata, 2)
			fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars + "*"
		elseif intLength = 8 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 3)
			strrightChars = Right(strtrimdata, 3)
			fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars + "*"
		elseif intlength = 9 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 3)
			strrightChars = Right(strtrimdata, 4)
			fmtparcelnum = strleftChars + "-" + strmidChars +  "-" + strrightChars
		elseif intlength = 10 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 3)
			strrightChars = Mid(strtrimdata, 6, 4)
			fmtparcelnum = strleftChars + "-" + strmidChars +  "-" + strrightChars
'Response.write(" the value of the fmtparcelnum in len 10 : " & fmtparcelnum)'
		End If
	End If

If varcid = 65  Then         'Creates  formatted parcel XX-XXXXX-XX  (2-5-4)
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
			strrightChars = Right(strtrimdata, 3)
			fmtparcelnum = strleftChars + "-" + strrightChars
	'		Response.write(" the value of the fmtparcelnum in len 5 : " & fmtparcelnum)
		elseif intlength = 6 Then
			strleftChars = Left(strtrimdata, 2)
			'strmidChars = Mid(strtrimdata, 3, 3)
			strrightChars = Right(strtrimdata, 4)
		'	fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
			fmtparcelnum = strleftChars + "-" + strrightChars
		'Response.write(" the value of the fmtparcelnum in len 6 : " & fmtparcelnum)'
		elseif intLength = 7 Then
			strleftChars = Left(strtrimdata, 2)
		'	strmidChars = Mid(strtrimdata, 3, 3)
			strrightChars = Right(strtrimdata, 5)
			fmtparcelnum = strleftChars + "-" + strrightChars
		elseif intLength = 8 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 5)
			strrightChars = Right(strtrimdata, 1)
			fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
		elseif intlength = 9 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 5)
			strrightChars = Right(strtrimdata, 2)
			fmtparcelnum = strleftChars + "-" + strmidChars +  "-" + strrightChars
		elseif intlength = 10 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 5)
			strrightChars = Mid(strtrimdata, 8, 2)
			fmtparcelnum = strleftChars + "-" + strmidChars +  "-" + strrightChars
		End If
	End If
If varcid = 42  Then         'Creates  formatted parcel XX-XXXXXX-X  (2-6-1)
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
			strrightChars = Right(strtrimdata, 3)
			fmtparcelnum = strleftChars + "-" + strrightChars
	'	Response.write(" the value of the fmtparcelnum in len 5 : " & fmtparcelnum)
		elseif intlength = 6 Then
			strleftChars = Left(strtrimdata, 2)
			'strmidChars = Mid(strtrimdata, 3, 3)
			strrightChars = Right(strtrimdata, 4)
		'	fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
			fmtparcelnum = strleftChars + "-" + strrightChars
		'Response.write(" the value of the fmtparcelnum in len 6 : " & fmtparcelnum)'
		elseif intLength = 7 Then
			strleftChars = Left(strtrimdata, 2)
		'	strmidChars = Mid(strtrimdata, 3, 3)
			strrightChars = Right(strtrimdata, 5)
			fmtparcelnum = strleftChars + "-" + strrightChars
		elseif intLength = 8 Then
			strleftChars = Left(strtrimdata, 2)
		'	strmidChars = Mid(strtrimdata, 3, 5)
			strrightChars = Right(strtrimdata, 6)
			fmtparcelnum = strleftChars + "-" + strrightChars
		elseif intlength = 9 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 6)
			strrightChars = Right(strtrimdata, 1)
			fmtparcelnum = strleftChars + "-" + strmidChars +  "-" + strrightChars
		elseif intlength = 10 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 6)
			strrightChars = Mid(strtrimdata, 9, 1)
			fmtparcelnum = strleftChars + "-" + strmidChars +  "-" + strrightChars
		End If
	End If

If varcid = 47 or varcid = 54 or varcid = 77   Then         'Creates  formatted parcel XX-XXXXXXX  (2-7)
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
			strrightChars = Right(strtrimdata, 3)
			fmtparcelnum = strleftChars + "-" + strrightChars
	'	Response.write(" the value of the fmtparcelnum in len 5 : " & fmtparcelnum)
		elseif intlength = 6 Then
			strleftChars = Left(strtrimdata, 2)
			'strmidChars = Mid(strtrimdata, 3, 3)
			strrightChars = Right(strtrimdata, 4)
		'	fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
			fmtparcelnum = strleftChars + "-" + strrightChars
		'Response.write(" the value of the fmtparcelnum in len 6 : " & fmtparcelnum)'
		elseif intLength = 7 Then
			strleftChars = Left(strtrimdata, 2)
		'	strmidChars = Mid(strtrimdata, 3, 3)
			strrightChars = Right(strtrimdata, 5)
			fmtparcelnum = strleftChars + "-" + strrightChars
		elseif intLength = 8 Then
			strleftChars = Left(strtrimdata, 2)
		'	strmidChars = Mid(strtrimdata, 3, 5)
			strrightChars = Right(strtrimdata, 6)
			fmtparcelnum = strleftChars + "-" + strrightChars
		elseif intlength = 9 Then
			strleftChars = Left(strtrimdata, 2)
		'	strmidChars = Mid(strtrimdata, 3, 6)
			strrightChars = Right(strtrimdata, 7)
			fmtparcelnum = strleftChars + "-" + strrightChars
		elseif intlength = 10 Then
			strleftChars = Left(strtrimdata, 2)
		'	strmidChars = Mid(strtrimdata, 3, 6)
			strrightChars = Mid(strtrimdata, 3, 7)
			fmtparcelnum = strleftChars + "-" + strrightChars
		End If
	End If



	Session("PrclNo") = fmtparcelnum
end Function
'*********
' The check for a hyphen function * * *    lem 9-16-08
Function ChckHyphen(strNumber, checkit)
	'dim checkit as string
	'dim strNumber as string
	'dim strFirst as string
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





If intParcelNo <> "" Then
'* * * * * * * * * * *'
 '  Handles the * error If someone enters just the *  this code should handle it.
if intParcelNo = "*"  then
'response.Redirect("SearchInput.asp")
'response.Write(" inside the parcelno = * ")
 '     Response.Clear
    Response.Redirect("disclaimer.asp?cid=" & varcid & "")
'response.Redirect("tax.asp")
else
'response.Write(" inside else the parcelno = * ")
End if
' * * * * * * * * * *
'** call the function if there is something to work with.

 	chkit2=ChckHyphen(intParcelNo, chkit)
 'Response.Write("the value of chkit: " & chkit )
 'Response.Write("the value of chkit2: " & chkit2 )
 'Response.Write(" the value 1 of intParcelNo is : " & intParcelNo  )
 	if chkit = 0 then
  	intParcelNo = fmtparcelnum(intParcelNo)
 	end if

	intParcelNo = replace(intParcelNo, "*", "%")
'Response.Write(" the value of intParcelNo is : " & intParcelNo  )
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL Like ' " & intParcelNo & "'"
	strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL Like '" & intParcelNo & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
	intParcelNo = replace(intParcelNo, "%", "*")
	intSet = 1
	strTitle1 = "Parcel No."
	strTitle2 = "Pay Yr"
	strTitle3 = "E911 Address"
	strTitle4 = "Name"
	strTitle5 = "Description"
End IF
If strAddress <> "" or strSAddress <> "" Then
	strAddress = replace(strAddress, "*", "%")
	strSAdress = replace(strSAdress, "*", "%")
	strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPADR1 Like '" & strAddress & "'"
	strAddress = replace(strAddress, "%", "*")
	strSAdress = replace(strSAdress, "%", "*")
	intSet = 2
	strTitle1 = "E911 Address"
	strTitle2 = "Parcel No."
	strTitle3 = "Pay Yr"
	strTitle4 = "Name"
	strTitle5 = "Description"
End IF
If strName <> "" Then
	strName = replace(strName, "*", "%")
	' If the multiple names on a parcel (owner, taxpayer, alt) are going to be able to search for then I must use the namesearch table.
	'strQueryString = "WHERE [NAMESEARCH].TXTNAM Like '" & strName & "' ORDER BY [NAMESEARCH].TXPRCL;"
	strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM Like '" & strName & "'"

	strName = replace(strName, "%", "*")
	intSet = 3
	strTitle1 = "Name"
	strTitle2 = "Parcel No."
	strTitle3 = "Pay Yr"
	strTitle4 = "Description"
End IF
If intTwp <> "" and intRange <> "" Then
	if intSect <> "" Then
	strQueryString = "WHERE  [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTOWN = " & intTwp & " AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXRANG = " & intRange & " AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXSECT = " & intSect & ""
	else
	strQueryString = "WHERE  [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTOWN = " & intTwp & " AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXRANG = " & intRange & ""
	end if
	intSet = 4
	strTitle1 = "Name"
	strTitle2 = "Parcel No."
	strTitle3 = "Pay Yr"
	strTitle4 = "Description"
	strTitle5 = ""
End IF


Set objCommand = Server.CreateObject("ADODB.Command")

'Fill in the command properties
objCommand.ActiveConnection = strConnect
if strName <> "" then
	objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
else
	if intTwp <> "" and intRange <> "" Then
		objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
	else
		objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
	end if
end if
'objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
objCommand.CommandType = 1

Set objRS = objCommand.Execute

Set objCommand = Nothing

Response.Write("<table width='1300'><th class='STitle'nowrap>" & strTitle1 & "</th><th class='STitle'>" & strTitle2 & "</th><th class='STitle'>" & strTitle3 & "</th><th class='STitle'>" & strTitle4 & "</th><th class='STitle'>" & strTitle5 & "</th>")
'intPrcl = objRS("TXPRCL")
'Response.Write("the session vars are  " & intPrcl & intPrclNo &  cid & Addr & Session("Nam") & Session("SecNam") & Session("TwnNam") & Session("RngNam") & strQueryString & " the end of it ")

While Not ObjRS.EOF
	Select Case intSet
	Case 1

			if intPrcl = objRS("TXPRCL") then
				'If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then
				'else
				'	Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel_Rock.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=1&yr=" & objRS("TXYEAR") & "'>" & objRS("TXPRCL") & "</a></td>")
				'	Response.Write("<td class='rText'>" & objRS("TXYEAR") & objRS("TXFLAG") & "</td>")
				'	Response.Write("<td class='rText'>" & objRS("TXPADR1") & "</td>")
				'	'Response.Write("<td class='rText'>" & objRS("TXFLAG") & "</td>")
				'	response.Write("<td class='rText'>" & objRS("TXTNAM") & "</td>")
				'	If  objRS("TXSECT") > 0 Then
				'		If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
				'			Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				'		else
				'			Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				'		end if
				'	Else
				'		If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
				'			Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				'		else
				'			Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				'		end if
				'	End If
	            '
				'	intPrcl = objRS("TXPRCL")
				'end If
			else
				Session("intYEAR") = objRS("TXYEAR")
				Session("intPYR") = (objRS("TXYEAR")-1)
				if objRS("TXFLAG")= "T" Then
				Response.Write("<tr><td class='rText'><a nowrap class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
				Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXPADR1") & "</td>")
				response.Write("<td class='rText'nowrap>" & objRS("TXTNAM") & "</td>")
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
				end if
				End If
				intPrcl = objRS("TXPRCL")
		End If



		'Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0  & "&cid=varcid'>" & objRS("TXPRCL") & "</a></td>")
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
		''Response.Write("<td class='rText'>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td></tr>")
	Case 2
		'Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0  & "&cid=varcid'>" & objRS("TXPADR1") & "</a></td>")
		'Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
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

		if intPrcl = objRS("TXPRCL")  then
					'Session("intYEAR") = objRS("TXYEAR")
					'If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then
					'else
					'	Response.Write("<tr><td class='rText'><a class='pLink'nowrap href='Parcel_Rock.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=1&yr=" & objRS("TXYEAR") & "'>" & objRS("TXPADR1") & "</a></td>")
					'	Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
					'	Response.Write("<td class='rText'>" & objRS("TXYEAR") & objRS("TXFLAG") & "</td>")
					'	Response.Write("<td class='rText'nowrap>" & objRS("TXTNAM") & "</td>")
					'	If  objRS("TXSECT") > 0 Then
					'		If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
					'			Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					'		else
					'			Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					'		end if
					'	Else
					'		If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
					'			Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					'		else
					'			Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					'		end if
					'	End If
					'	intPrcl = objRS("TXPRCL")
					'End If
				else
					Session("intYEAR") = objRS("TXYEAR")
					Session("intPYR") = (objRS("TXYEAR")-1)
					if objRS("TXFLAG")= "T" Then
					Response.Write("<tr><td class='rText'><a class='pLink'nowrap href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPADR1") & "</a></td>")
					Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
					Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
					Response.Write("<td class='rText'nowrap>" & objRS("TXTNAM") & "</td>")
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
					end if
					intPrcl = objRS("TXPRCL")
				End If



	Case 3


        if intPrcl = objRS("TXPRCL") then
				''Session("intYEAR") = objRS("TXYEAR")
			'If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then
			'Response.write(" this is a test for year and flag = T ")
			'else
			'	Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel_Rock.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=1&yr=" & objRS("TXYEAR") & "'>" & objRS("TXTNAM") & " </a></td>")
			'	'Response.Write("<td class='rText'>" & objRS("TXTNAM") &  "</td>")
			'	Response.Write("<td class='rText'>" & objRS("TXPRCL") & "</td>")
			'	Response.Write("<td class='rText'>" & objRS("TXYEAR") & objRS("TXFLAG") & "</td>")
			'	If  objRS("TXSECT") > 0 Then
			'		If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
			'			Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			'		else
			'			Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			'		end if
			'	Else
			'		If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
			'			Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			'		else
			'			Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			'		end if
			'	End If
			'	intPrcl = objRS("TXPRCL")
			'End If
		else
				Session("intYEAR") = objRS("TXYEAR")
				Session("intPYR") = (objRS("TXYEAR")-1)
				if objRS("TXFLAG")= "T" Then
				Response.Write("<tr><td class='rText'><a class='pLink'nowrap href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXTNAM") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
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
				end if
				intPrcl = objRS("TXPRCL")

		End If


		'Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0'>" & objRS("TXTNAM") & "</a></td>")
		'Response.Write("<td class='rText'>" & objRS("TXPRCL") & "</td>")
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
	Case 4
		if intPrcl = objRS("TXPRCL") then
		else
			'Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0'>" & objRS("TXPRCL") & "</a></td>")
			'Response.Write("<td class='rText'>" & objRS("TXPADR1") & "</td>")
			'Response.Write("<td class='rText'>" & objRS("TXYEAR") & "</td>")
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
				Session("intYEAR") = objRS("TXYEAR")
				Session("intPYR") = (objRS("TXYEAR")-1)
				if objRS("TXFLAG")= "T" Then
					Response.Write("<tr><td class='rText'><a class='pLink'nowrap href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXTNAM") & "</a></td>")
					Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
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
				end if
					intPrcl = objRS("TXPRCL")

		End If

		'Response.Write("<td class='rText'>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") & " " & objRS("TXDSC1") & "</td></tr>")
	End Select
	objRS.MoveNext
Wend

objRS.Close
Set objRS = Nothing
%>

</table>
</body>
</html>
