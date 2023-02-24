<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Parcel Search Results</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
	Dim varcid
	varcid = Request.Form("cid")
	response.Write("<link rel='stylesheet' href='" & varcid & ".css' type='text/css'>")
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
strAddress = Request.Form("pAddress")
strSAddress = Request.Form("sAddress")
strName = Request.Form("fName")
intSect = Request.Form("sName")
intTwp = Request.Form("tName")
intRange = Request.Form("rName")
Session("RngNam") = intRange
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

If varcid = 34 or varcid=48 or varcid=51 or varcid=64 or varcid=74 or varcid=84 or varcid=87  Then         'Creates  formatted parcel XX-XXX-XXXX  (2-3-4)
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
'			Response.write(" the value of the fmtparcelnum in len 3 : " & fmtparcelnum)
		elseif intlength = 4 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Right(strtrimdata, 2)
			fmtparcelnum = strleftChars + "-" + strmidChars
		'Response.write(" the value of the fmtparcelnum in len 4 : " & fmtparcelnum)
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



If strParcelNo <> "" Then
'* * * * * * * * * * *'
 '  Handles the * error If someone enters just the *  this code should handle it.
if strParcelNo = "*"  then
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
'Response.Write(" the value of strParcelNo is : " & strParcelNo  )
'** call the function if there is something to work with.
'Response.Write(" the value 1 of strParcelNo is : " & strParcelNo  )
 	chkit2=ChckHyphen(strParcelNo, chkit)
' Response.Write("the value of chkit: " & chkit )
' Response.Write("the value of chkit2: " & chkit2 )
' Response.Write(" the value 2 of strParcelNo is : " & strParcelNo  )
 	if chkit = 0 then
  	strParcelNo = fmtparcelnum(strParcelNo)
 	end if
'Response.Write(" the value23 of strParcelNo is : " & strParcelNo  )
	strParcelNo = replace(strParcelNo, "*", "%")
	'Response.Write(" the second value of : " & strParcelNo)
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL Like '" & strParcelNo & "'"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL Like '" & strParcelNo & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR;"
	strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL Like '" & strParcelNo & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"

	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = '" & strParcelNo & "';"
	strParcelNo = replace(strParcelNo, "%", "*")
	intSet = 1
	strHdgNote="Left click on the desired parcel to view additional information."
	strTitle1 = "Parcel No."
	strTitle2 = "Year"
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
	strHdgNote="Left click on the desired address to view additional information."
	strTitle1 = "E911 Address"
	strTitle2 = "Parcel No."
	strTitle3 = "Year"
	'strTitle4 = "Value/Tax"
	strTitle4 = "Name"
	strTitle5 = "Description"
End IF
If strName <> "" Then
	newSearch = ChckStar(strName, checkstaradd)
	strName = checkstaradd
	strName = replace(strName, "*", "%")
	strName = replace(strName, "'", "''")
	'objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] INNER JOIN NameSearch ON [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL = NameSearch.TXPRCL WHERE [NAMESEARCH].TXTNAM Like '" & strName & "'"
	'objCommand.CommandType = 1
	'Set objRS = objCommand.Execute
	'Set objCommand = Nothing
	'strQueryString = "WHERE [NAMESEARCH].TXTNAM Like '" & strName & "'"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM Like '" & strName & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL;"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM Like '" & strName & "' OR [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXANAM Like '" & strName & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR;"
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM Like '" & strName & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
    strQueryString = " INNER JOIN [Table 1 - Name/Addr/Desc/Tax/Recap Info] ON [Namesearch].TXPRCL = [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL AND [Namesearch].TXYEAR = [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR AND [Namesearch].TXFLAG = [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG WHERE  [Namesearch].TXTNAM Like '" & strName & "' ORDER BY [Namesearch].TXTNAM;"
	'strQueryString = "WHERE [NAMEQUERY].TXTNAM Like '" & strName & "' ORDER BY [NAMEQUERY].TXPRCL;"
	'strQueryString = "WHERE [NAMESEARCH].TXTNAM Like '" & strName & "' ORDER BY [NAMESEARCH].TXPRCL;"
	'The line below is the good to go previous SQL"""""""""""""""""
	'strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTNAM Like '" & strName & "'"
	strName = replace(strName, "''", "'")
	strName = replace(strName, "%", "*")
	intSet = 3
	strHdgNote="Left click on the desired name to view additional information."
	strTitle1 = "Name"
	strTitle2 = "Parcel No."
	strTitle3 = "Year"
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
	strHdgNote="Left click on the desired parcel to view additional information."
	strTitle1 = "Parcel No."
	strTitle2 = "Year"
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
	'objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
	objCommand.CommandText = "SELECT [Namesearch].*, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXDSC1, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXBLOK, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXSECT, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPLATD, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXLOT, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXTOWN, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXRANG FROM [Namesearch] " & strQueryString
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
'Response.Write(objCommand.CommandText)
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


Response.Write("<font color='red'><b>" & strHdgNote & "</b></font>")
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

		If objRS("TXFLAG") = "V" then
			strFlag21 = " Value for Tax Payable "
		end if
	Select Case intSet

	Case 1
		if intPrcl = objRS("TXPRCL") then
			'If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then
			If  objRS("TXFLAG")= "T" then
			Session("intTAXYEAR") = objRS("TXYEAR")
			'If Session("intYEAR") = objRS("TXYEAR")  then
			'else
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
				'Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=1&yr=" & objRS("TXYEAR") & "'>" & objRS("TXPRCL") & "</a></td>")

					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag & "</td>")
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
				'Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
				else
					If varcid = 61 then
					'Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else
					If varcid = 84 then
					'Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
					else

					'Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag &"</td>")
					end if
					end if
				end if
				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
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
			Session("intTAXYEAR") = objRS("TXYEAR")
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

				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
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

			intPrcl = objRS("TXPRCL")
	End If
	Case 2

		if intPrcl = objRS("TXPRCL")  then
			'Session("intYEAR") = objRS("TXYEAR")
			'If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then
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


                Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
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


				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
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


'			Session("intYEAR") = objRS("TXYEAR")
'			Session("intPYR") = (objRS("TXYEAR")-1)
'			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0'>" & objRS("TXPADR1") & "</a></td>")
'			Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
'			Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & "</td>")
'			Response.Write("<td class='rText'nowrap>" & objRS("TXTNAM") & "</td>")
'			If  objRS("TXSECT") > 0 Then
'				If (objRS("TXLOT") = 0) or (objRS("TXBLOK") = 0) Then
'					Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
'				else
'					Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
'				end if
'			Else
'				If (objRS("TXLOT") > 0) or (objRS("TXBLOK") > 0) Then
'					Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
'				else
'					Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
'				end if
'			End If
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
			'If Session("intYEAR") = objRS("TXYEAR") OR objRS("TXFLAG")= "T" then
			'If Session("intYEAR") = objRS("TXYEAR")  then
			If  objRS("TXFLAG")= "T" then
			Session("intTAXYEAR") = objRS("TXYEAR")
			'else
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXTNAM") & "</a></td>")
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
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=0&yr=" & objRS("TXYEAR") & "&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXTNAM")  &" </a></td>")
					'Response.Write("<td class='rText'nowrap>" & objRS("TXTNAM") &  "</td>")
					Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")


					Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
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
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXTNAM")  & "</a></td>")
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
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&rid=0&yr=" & objRS("TXYEAR") & "&cid=" & varcid & "&varintParcelNo=" & strParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXTNAM") & " </a></td>")
				'Response.Write("<td class='rText'nowrap>" & objRS("TXTNAM") &  "</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")


				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
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

				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
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


				Response.Write("<td class='rText'nowrap>" & objRS("TXYEAR") & strFlag21 & (objRS("TXYEAR") + 1) & "</td>")
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

'Session("intREC") = objRSV2("RecNum")
%>

</table>
</body>
</html>
