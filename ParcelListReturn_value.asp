<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Parcel Search Results</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
	Dim varcid
	'cid = Session("CountyID")
	'varcid = Request.Form("cid")
	varcid = request.QueryString("CID")
	response.Write("<link rel='stylesheet' href='" & varcid & ".css' type='text/css'>")
	'response.Write (" variable varcid :" & varcid  )
%>
</head>
<!-- #include file="insDB.asp" -->
<body>
<%
If cName = "Douglas" Then
	Response.Write("<input name='returnHomeButton' type='button' value='Return to Douglas Home Page' onClick=window.location='http://www.co.douglas.mn.us'>")
End If

strParcelNo = Request.QueryString("varintParcelNo")
strAddress = request.QueryString("varstrAddress")
strSAddress = request.QueryString("sAddress")
strName = request.QueryString("varstrName")
intSect = request.QueryString("varintSect")
intTwp = request.QueryString("varintTwp")
intRange = request.QueryString("varintRange")
'Session("RngNam") = intRange
intRC = Session("intREC")

'***  IF user enters a * in the search box the user is taken back to the disclaimer page. added 9-23-08 LEM
if strAddress = "*" or strSAddress = "*" then
    Response.Redirect("disclaimer.asp?cid=" & varcid & "")
End If
If strName = "*" Then
    Response.Redirect("disclaimer.asp?cid=" & varcid & "")
End If


'**  Function to format the entry of the parcel number search from the user
'**  this will allow the user to enter numbers only in the text field of the form.
'**  this is in ND web tax' Added on 9-4-08 LEM
Function fmtparcelnum(strData)
	Dim strtrimdata
	strData = Request.Form("ParcelNo")
'response.write("the value of ParcelNo : " & strParcelNo )
	strtrimdata = Trim(strData)
	intLength = Len(strtrimdata)

If varcid = 21 or varcid = 26 or varcid=41 or varcid=45 or varcid=53 or varcid=61 or varcid=67 or varcid=75 or varcid=76  Then         'Creates  formatted parcel XX-XXXX-XXX  (2-4-3)
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
'Response.write(" the value of the fmtparcelnum in len 4 : " & fmtparcelnum)
		elseif intlength = 5 Then
			strleftChars = Left(strtrimdata, 2)
		'	strmidChars = Mid(strtrimdata, 3, 2)
			strrightChars = Right(strtrimdata, 3)
		'	fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
			fmtparcelnum = strleftChars + "-" + strrightChars
'Response.write(" the value of the fmtparcelnum in len 5 : " & fmtparcelnum)
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
		End If
	End If

If varcid = 34 or varcid=48 or varcid=51 or varcid=64 or varcid=84 or varcid=74 or varcid=87  Then         'Creates  formatted parcel XX-XXX-XXXX  (2-3-4)
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
		'Response.write(" the value of the fmtparcelnum in len 5 : " & fmtparcelnum)
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
		End If
	End If

If varcid = 65  Then         'Creates  formatted parcel XX-XXXXX-XX  (2-5-2)
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
		'Response.write(" the value of the fmtparcelnum in len 5 : " & fmtparcelnum)
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
		End If
	End If
If varcid = 42  Then         'Creates  formatted parcel XX-XXXXXX-X  (2-6-1)
'Response.write(" the value of the fmtparcelnum in varcid = 42 : " & fmtparcelnum)
	If intLength = 1 Then
			fmtparcelnum = strtrimdata
		elseif intlength = 2 Then
			fmtparcelnum = strtrimdata
		elseif intlength = 3 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Right(strtrimdata, 1)
			fmtparcelnum = strleftChars + "-" + strmidChars
			Response.write(" the value of the fmtparcelnum in len 3 : " & fmtparcelnum)
		elseif intlength = 4 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Right(strtrimdata, 2)
			fmtparcelnum = strleftChars + "-" + strmidChars
		'Response.write(" the value of the fmtparcelnum in len 4 : " & fmtparcelnum)
		elseif intlength = 5 Then
			strleftChars = Left(strtrimdata, 2)
			strrightChars = Right(strtrimdata, 3)
			fmtparcelnum = strleftChars + "-" + strrightChars
		Response.write(" the value of the fmtparcelnum in len 5 : " & fmtparcelnum)
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
		End If
	End If
If varcid = 47 or varcid = 54 or varcid = 77 or varcid = 78  Then         'Creates  formatted parcel XX-XXXXXXX  (2-7)
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
	'Response.Write(" the value of strNumber inside chdkhyphen func is : " & strNumber  )
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


If cName = "Douglas" Then
	Response.Write("<input name='returnHomeButton' type='button' value='Return to Douglas Home Page' onClick=window.location='http://www.co.douglas.mn.us'>")
End If



If strParcelNo <> "" Then
	strParcelNo = replace(strParcelNo, "*", "%")
	'Response.Write(" the second value of : " & strParcelNo)
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL Like '" & strParcelNo & "'"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL Like '" & strParcelNo & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR;"
	strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL Like '" & strParcelNo & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL Like '" & strParcelNo & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strParcelNo & "' AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR = 2006;"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strParcelNo & "';"
	strParcelNo = replace(strParcelNo, "%", "*")
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
'	strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPADR1 Like '" & strAddress & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
	strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPADR1 Like '" & strAddress & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPADR1, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
	strAddress = replace(strAddress, "%", "*")
	strSAdress = replace(strSAdress, "%", "*")

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
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM Like '" & strName & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL;"
	'strQueryString = "WHERE [NAMEQUERY].TXTNAM Like '" & strName & "' ORDER BY [NAMEQUERY].TXPRCL;"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM Like '" & strName & "'"
	strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM Like '" & strName & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
	strName = replace(strName, "%", "*")
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


Set objCommand = Server.CreateObject("ADODB.Command")

'Fill in the command properties
objCommand.ActiveConnection = strConnect
if strName <> "" then
	'objCommand.CommandText = "SELECT * FROM [NameQuery] " & strQueryString
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
'Response.Write("the query 1 : " & strQueryString )
Set objRS = objCommand.Execute

Set objCommand = Nothing

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
Session("intTAXYEAR") = objRS("TXYEAR")
	If objRS("TXFLAG") = "T" then
		strFlag = " Tax"
	end if
	If objRS("TXFLAG") = "V" then
		strFlag = " Value"
	end if

	If varcid=21 then
		If objRS("TXFLAG") = "V" then
			strFlag21 = " Value for Tax Payable "
		end if
	end if
	If varcid= 61 then
		If objRS("TXFLAG") = "V" then
			strFlag21 = " Value for Tax Payable "
		end if
	end if
	Select Case intSet
	Case 1

		if intPrcl = objRS("TXPRCL") then
			'Session("intYEAR") = objRS("TXYEAR")
						'If Session("intYEAR") = objRS("TXYEAR")  then
						If  objRS("TXFLAG")= "T" then
						Session("intTAXYEAR") = objRS("TXYEAR")
						'else
							Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
							'Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=1&yr=" & objRS("TXYEAR") & "'>" & objRS("TXPRCL") & "</a></td>")
					If varcid = 21 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					If varcid = 61 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
					end if
					end if

'							Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
							Response.Write("<td class='rText'nowrap>" & objRS("TXPADR1") & "</td>")
							'Response.Write("<td class='rText'>" & objRS("TXFLAG") & "</td>")
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
							End If

							intPrcl = objRS("TXPRCL")
						end If
			'** added code to get  the 'V' for the next year.

						'If (Session("intYEAR")+ 1) = objRS("TXYEAR") and objRS("TXFLAG")= "V" then
						If  objRS("TXFLAG")= "V" then
						'else
							Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=0&yr=" & objRS("TXYEAR") & "&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
					If varcid = 21 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					If varcid = 61 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
					end if
					end if

'							Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
							Response.Write("<td class='rText'nowrap>" & objRS("TXPADR1") & "</td>")
							'Response.Write("<td class='rText'>" & objRS("TXFLAG") & "</td>")
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
							End If

							intPrcl = objRS("TXPRCL")
						end If

			'** end the added code to get  the 'V' for the next year.


		else


' The added code here for an error in Nobles County
' this is to correct a new parcel split done a few days ago that fails to
'  show the new parcel correctly. I have to get the Parcel_value.asp to be the new LINK so that I go
'  to the 'Value' side of the information.

			'If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then
			If  objRS("TXFLAG")= "T" then
' added code here to work through the Nobles County problem.
			Session("intYEAR") = objRS("TXYEAR")
			Session("intPYR") = (objRS("TXYEAR")-1)
			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
			Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
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
			End If

' end the new added code here



			else
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=0&yr=" & objRS("TXYEAR") & "&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
					If varcid = 21 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					If varcid = 61 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
					end if
					end if

'				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag & "</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXPADR1") & "</td>")
				'Response.Write("<td class='rText'>" & objRS("TXFLAG") & "</td>")
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
				End If

				intPrcl = objRS("TXPRCL")
			end If
'  end the new added code here.


			intPrcl = objRS("TXPRCL")
		End If

	Case 2
		if intPrcl = objRS("TXPRCL")  then
			'Session("intYEAR") = objRS("TXYEAR")
			'If Session("intYEAR") = objRS("TXYEAR")  then
			If  objRS("TXFLAG")= "T" then
			Session("intTAXYEAR") = objRS("TXYEAR")
			'else
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPADR1") & "</a></td>")
				'Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=1&yr=" & objRS("TXYEAR") & "'>" & objRS("TXPADR1") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
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
				intPrcl = objRS("TXPRCL")
			End If

'** added code to get  the 'V' for the next year.

			'If (Session("intYEAR")+ 1) = objRS("TXYEAR") and objRS("TXFLAG")= "V" then
			If objRS("TXFLAG")= "V" then
			'else
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=0&yr=" & objRS("TXYEAR") & "&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPADR1") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")

					If varcid = 21 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					If varcid = 61 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
					end if
					end if
'				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
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
				intPrcl = objRS("TXPRCL")
			end If

'** end the added code to get  the 'V' for the next year.

		else

'added code here for a split parcel done in the last year and there is only the new 'Value' record
' to show to user.
			'If Session("intYEAR") = objRS("TXYEAR") or objRS("TXFLAG")= "T" then
			If  objRS("TXFLAG")= "T" then
			Session("intTAXYEAR") = objRS("TXYEAR")
' last added code here

			Session("intYEAR") = objRS("TXYEAR")
			Session("intPYR") = (objRS("TXYEAR")-1)
			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPADR1") & "</a></td>")
			Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
			Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
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

' end last added code here.


				else
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=0&yr=" & objRS("TXYEAR") & "&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPADR1") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")

					If varcid = 21 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					If varcid = 61 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
					end if
					end if
'				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
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
				intPrcl = objRS("TXPRCL")
			End If

' added code ended here.

			intPrcl = objRS("TXPRCL")
		End If

	Case 3

		if intPrcl = objRS("TXPRCL") then
				'Session("intYEAR") = objRS("TXYEAR")
			'If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then
			'If Session("intYEAR") = objRS("TXYEAR")  then
			If  objRS("TXFLAG")= "T" then
			Session("intTAXYEAR") = objRS("TXYEAR")
			'else
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXTNAM") & objRS("TXANAM") & "</a></td>")
				'Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=1&yr=" & objRS("TXYEAR") & "'>" & objRS("TXTNAM") & objRS("TXANAM") &" </a></td>")
				'Response.Write("<td class='rText'nowrap>" & objRS("TXTNAM") &  "</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
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

'** added code to get  the 'V' for the next year.

			'If (Session("intYEAR")+ 1) = objRS("TXYEAR") and objRS("TXFLAG")= "V" then
			If  objRS("TXFLAG")= "V" then
			'else
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=0&yr=" & objRS("TXYEAR") & "&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXTNAM") & objRS("TXANAM") &" </a></td>")
					'Response.Write("<td class='rText'nowrap>" & objRS("TXTNAM") &  "</td>")
					Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")

					If varcid = 21 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					If varcid = 61 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
					end if
					end if
'					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
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

'** end the added code to get  the 'V' for the next year.
		else

			'If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then

'special little exception here . When there is no past year record in this If Else Endif  conditional statement
' then the next record is the same as the last record  ( parcel # )  then the IF sends it to the else
' this corrects the outcome when this rare event happens.     1-31-07
			If  objRS("TXFLAG")= "T" then
				Session("intTAXYEAR") = objRS("TXYEAR")
				Session("intYEAR") = objRS("TXYEAR")
				Session("intPYR") = (objRS("TXYEAR")-1)
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXTNAM") & objRS("TXANAM") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
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

			else
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=0&yr=" & objRS("TXYEAR") & "&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXTNAM") & objRS("TXANAM") &" </a></td>")
				'Response.Write("<td class='rText'nowrap>" & objRS("TXTNAM") &  "</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")

					If varcid = 21 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					If varcid = 61 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
					end if
					end if
'				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag & "</td>")
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

		End If

	Case 4
			'Response.Write("<tr><td class='rText'><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0'>" & objRS("TXPRCL") & "</a></td>")
			'Response.Write("<td class='rText'>" & objRS("TXYEAR") & objRS("TXFLAG") & "</td>")
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
				'If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then
				'If Session("intYEAR") = objRS("TXYEAR")  then
				If objRS("TXFLAG")= "T" then
					Session("intTAXYEAR") = objRS("TXYEAR")
				'else
						Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
						'Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=1&yr=" & objRS("TXYEAR") & "'>" & objRS("TXPRCL") & "</a></td>")
						Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
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
						End If
						intPrcl = objRS("TXPRCL")
				End If

'** added code to get  the 'V' for the next year.

			'If (Session("intYEAR")+ 1) = objRS("TXYEAR") and objRS("TXFLAG")= "V" then
			If  objRS("TXFLAG")= "V" then
			'else
			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=0&yr=" & objRS("TXYEAR") & "&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")

					If varcid = 21 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					If varcid = 61 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
					end if
					end if
'							Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
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
							End If
			intPrcl = objRS("TXPRCL")
			end If

'** end the added code to get  the 'V' for the next year.

			else
			If  objRS("TXFLAG")= "T" then
' added code here to work through the Nobles County problem.
			Session("intYEAR") = objRS("TXYEAR")
					Session("intYEAR") = objRS("TXYEAR")
					Session("intPYR") = (objRS("TXYEAR")-1)
					Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
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
					End If

			else

				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=0&yr=" & objRS("TXYEAR") & "&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")

					If varcid = 21 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					If varcid = 61 then
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
					end if
					end if
'				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXPADR1") & "</td>")
				'Response.Write("<td class='rText'>" & objRS("TXFLAG") & "</td>")
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
				End If

					intPrcl = objRS("TXPRCL")
			End If
		End if
		'Response.Write("<td class='rText'>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") & " " & objRS("TXDSC1") & "</td></tr>")
	End Select
	objRS.MoveNext
	'If objRS.EOF  then
	'Response.Write("End Of File")
	'end if
Wend

objRS.Close
Set objRS = Nothing
%>

</table>
</body>
</html>
