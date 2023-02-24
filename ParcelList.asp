<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Parcel Search Results</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
	Dim cid
	cid = Request.Form("cid")

	response.Write("<link rel='stylesheet' href='" & cid & ".css' type='text/css'>")
%>
</head>
<!-- #include file="insDB.asp" -->
<body>
<%

Dim objCommand, objCommand2, objRS, objRSV, objRSV2, strParcelNo, intPrcl, intRC,  strAddress, Addr, strQueryString, intSet, strTitle1, strTitle2, strTitle3, strTitle4, strTitle5, strName, intPrclNo, strAddr, strNam
strParcelNo = Request.Form("ParcelNo")
strAddress = Request.Form("pAddress")
strSAddress = Request.Form("sAddress")
strName = Request.Form("ppName")
intParcelNo = Request.Form("ParcelNo")
intSect = Request.Form("sName")
intTwp = Request.Form("tName")
intRange = Request.Form("rName")
strPlat = Request.Form("pName")
intBlock = Request.Form("bName")
strLot = Request.Form("lName")

'***  IF user enters a * in the search box the user is taken back to the disclaimer page.'
if strAddress = "*" or strSAddress = "*" or strName = "*" then
    Response.Redirect("disclaimer.asp?cid=" & cid & "")
End If


Function fmtparcelnum(strData)
	Dim strtrimdata
	strData = Request.Form("ParcelNo")
	strtrimdata = Trim(strData)
	intLength = Len(strtrimdata)
If cid = 47 or cid = 30 or cid = 31 or cid = 34 or cid = 37 or cid = 23 or cid = 41 or cid = "02" Then
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
	elseif intlength = 10 Then
			strleftChars = Left(strtrimdata, 2)
			strrightChars = Mid(strtrimdata, 3, 7)
			fmtparcelnum = strleftChars + "-" + strrightChars
	End If
End If
If cid = 27  Then
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
		elseif intlength = 6 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 2)
			strrightChars = Right(strtrimdata, 2)
			fmtparcelnum = strleftChars + "-" + strmidChars + "-" + strrightChars
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
		elseif intlength = 10 Then
			strleftChars = Left(strtrimdata, 2)
			strmidChars = Mid(strtrimdata, 3, 2)
			strrightChars = Mid(strtrimdata, 5, 5)
			fmtparcelnum = strleftChars + "-" + strmidChars +  "-" + strrightChars
		End If
	End If

If cid = 13  Then
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
		elseif intlength = 6 Then
			strleftChars = Left(strtrimdata, 2)
			strrightChars = Right(strtrimdata, 4)
			fmtparcelnum = strleftChars + "-" + strrightChars
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

end Function


Function ChckStar(strSearch, checkstar)
	starwhere = InStr(strSearch, "*")
	strLen = Len(strSearch)

	if starwhere = 0 Then
	checkstar = strSearch + "*"
	else

	checkstar = strSearch
	strStarLoc = Mid(strSearch, strLen, 1)
	End If
End Function

Function ChckHyphen(strNumber, checkit)
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

If strParcelNo <> "" Then

if strParcelNo = "*"  then
    Response.Redirect("SearchInput.asp?cid=" & cid & "")
End if

	chkit2=ChckHyphen(strParcelNo, chkit)
	if chkit = 0 then
 	strParcelNo = fmtparcelnum(strParcelNo)
 	end if
	NewParcl =  ChckStar(strParcelNo, checkstr)
	strParcelNo = checkstr

	strParcelNo = replace(strParcelNo, "*", "%")
	strQueryString = "WHERE TXPRCL Like '" & strParcelNo & "' ORDER BY TXPRCL, TXYEAR, TXFLAG;"
	intParcelNo = replace(intParcelNo, "%", "*")
	intSet = 1
	strTitle1 = "Parcel No."
	strTitle2 = "Year"
	strTitle3 = "E911 Address"
	strTitle4 = "Name"
	strTitle5 = "Description"
End IF

If strAddress <> "" or strSAddress <> "" Then
	strAddress = replace(strAddress, "*", "%")
	strSAdress = replace(strSAdress, "*", "%")
'	strQueryString = "WHERE PROPADR Like '" & strAddress & "'"
	strQueryString = "WHERE PROPADR Like '" & strAddress & "' ORDER BY PROPADR, TXYEAR, TXFLAG;"
	strAddress = replace(strAddress, "%", "*")
	strSAdress = replace(strSAdress, "%", "*")
	intSet = 2
	strTitle1 = "E911 Address"
	strTitle2 = "Parcel No."
	strTitle3 = "Year"
	strTitle4 = "Name"
	strTitle5 = "Description"
End IF

If strName <> "" Then
	strName = replace(strName, "*", "%")
	strName = replace(strName, "'", "''")
	 strQueryString = "WHERE TPNAME Like '" & strName & "' ORDER BY TPNAME, TXPRCL, TXFLAG;"
	 strName = replace(strName, "''", "'")
	strName = replace(strName, "%", "*")
	intSet = 3
	strTitle1 = "Name"
	strTitle2 = "Parcel No."
	strTitle3 = "Year"
	strTitle4 = "Description"
End IF

If intTwp <> "" and intRange <> "" Then
	if intSect <> "" Then

	strQueryString = "WHERE TXTOWN = " & intTwp & " AND TXRANG = " & intRange & " AND TXSECT = " & intSect & " ORDER BY TXPRCL, TXYEAR, TXFLAG;"
	else

	strQueryString = "WHERE TXTOWN = " & intTwp & " AND TXRANG = " & intRange & " ORDER BY TXPRCL, TXYEAR, TXFLAG;"
	end if
	intSet = 4
	strTitle1 = "Parcel No."
	strTitle2 = "Year"
	strTitle3 = "E911 Address"
	strTitle4 = "Name"
	strTitle5 = "Description"
End IF

' Start the Plat Block and Lot queries here *******************************************************



If strPlat <> "" Then
	newSearch = ChckStar(strPlat, checkstarplat)
	strPlat = checkstarplat
	strPlat = replace(strPlat, "*", "%")
	strQueryString = "WHERE TXPLATD Like '" & strPlat & "'"
	strPlat = replace(strPlat, "%", "*")
	intSet = 5
End If
If intBlock <> "" Then
	strQueryString = "WHERE TXBLOK = " & intBlock & ""
	intSet = 5
End If
If strLot <> "" Then
	strQueryString = "WHERE TXLOT = '" & strLot & "'"
	intSet = 5
End If
If intBlock <> "" and strLot <> "" Then
	strQueryString = "WHERE TXBLOK = " & intBlock & " AND TXLOT = '" & strLot & "'"
	intSet = 5
End If
If strPlat <> "" and strLot <> ""  Then
	strPlat = replace(strPlat, "*", "%")
	strQueryString = "WHERE TXPLATD Like '" & strPlat & "' AND TXLOT = '" & strLot & "'"
	strPlat = replace(strPlat, "%", "*")
	intSet = 5
End If
If strPlat <> "" and intBlock <> ""  Then
	strPlat = replace(strPlat, "*", "%")
	strQueryString = "WHERE TXPLATD Like '" & strPlat & "' AND TXBLOK = " & intBlock & ""
	strPlat = replace(strPlat, "%", "*")
	intSet = 5
End If

If strPlat <> "" and intBlock <> "" and strLot <> "" Then
	strPlat = replace(strPlat, "*", "%")
	strQueryString = "WHERE TXPLATD Like '" & strPlat & "' AND TXLOT = '" & strLot & "' AND TXBLOK = " & intBlock & ""
	strPlat = replace(strPlat, "%", "*")
	intSet = 5
End If

If intSet = 5 Then
	strTitle1 = "Parcel No."
	strTitle2 = "E911 Address"
	strTitle3 = "Year"
	strTitle4 = "Name"
	strTitle5 = "Description"
End If


Set objCommand = Server.CreateObject("ADODB.Command")

objCommand.ActiveConnection = strConnect

	Select Case intSet
	Case 1
		objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
	Case 2
		objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
	Case 3
		objCommand.CommandText = "SELECT * FROM [NAMESEARCH] " & strQueryString
	Case 4
		objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
	Case 5
		objCommand.CommandText = "SELECT * FROM [Table 1 - Name/Addr/Desc/Tax/Recap Info] " & strQueryString
	End Select

objCommand.CommandType = 1

Set objRS = objCommand.Execute

Set objCommand = Nothing

'* * *   Create the RecordSets    Below   * * * * * * * * * *

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE PARCEL = '" & strQueryString & "' ORDER BY PARCEL, YEAR;"

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRSV = objCommand.Execute
Set objCommand = Nothing

Set objCommand = Server.CreateObject("ADODB.Command")

strPID = request.QueryString("pid")
strTID = request.QueryString("tid")
strQueryString = request.QueryString("pid")
strQueryString = "WHERE PARCEL = '" & strQueryString & "' ORDER BY Parcel, Year DESC, RecNum DESC;"

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 9 - Value Info] " & strQueryString
objCommand.CommandType = 1

Set objRSV2 = objCommand.Execute

Set objCommand = Nothing


Response.Write("<table width='1200'><th class='STitle'nowrap>" & strTitle1 & "</th><th class='STitle'>" & strTitle2 & "</th><th class='STitle'>" & strTitle3 & "</th><th class='STitle'>" & strTitle4 & "</th><th class='STitle'>" & strTitle5 & "</th>")

While Not objRS.EOF

Response.Flush

	If objRS("TXFLAG") = "T" then
		strFlag = " Tax"

			strFlag21 = objRS("TXYEAR") & " Tax"
	end if
	If objRS("TXFLAG") = "V" then
		strFlag = " Value"
		If cid=30 then
			strFlag21 = objRS("TXYEAR") & " Specials"
		Else
			strFlag21 = objRS("TXYEAR") & " Value"
		end if
	end if

	Select Case intSet
	Case 1														'Search by Parcel
		If  objRS("TXFLAG")= "T" then
			Response.Write("<tr><td class='rText'><a nowrap class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
		Else
			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=0&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
		End if
		Response.Write("<td class='rText'nowrap width='100'>"& strFlag21 &"</td>")
		Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
		response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
		If  objRS("TXSECT") > 0 Then
			If (objRS("TXLOT") = "") or (objRS("TXBLOK") = 0) Then
				Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			else
				Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			end if
		Else
			If (objRS("TXLOT") <> "") or (objRS("TXBLOK") > 0) Then
				Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			else
				Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			end if
		End If
		intPrcl = objRS("TXPRCL")

	Case 2														'Search by Address
			If  objRS("TXFLAG") = "T" then
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=1&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("PROPADR") & "</a></td>")
			Else
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=0&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("PROPADR") & "</a></td>")
			End if
			Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
			Response.Write("<td class='rText'nowrap width='100'>"& strFlag21 &"</td>")
			Response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
			If  objRS("TXSECT") > 0 Then
				If (objRS("TXLOT") = "") or (objRS("TXBLOK") = 0) Then
					Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				else
					Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				end if
			Else
				If (objRS("TXLOT") <> "") or (objRS("TXBLOK") > 0) Then
					Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				else
					Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				end if
			End If
			intPrcl = objRS("TXPRCL")

	Case 3														'Search by Name
		If  objRS("TXFLAG") = "T" then
			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=1&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TPNAME") & "</a></td>")
		Else
			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=0&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "'>" & objRS("TPNAME") & "</a></td>")
		End if
		Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
		Response.Write("<td class='rText'nowrap width='100'>"& strFlag21 &"</td>")
		If  objRS("TXSECT") > 0 Then
			If (objRS("TXLOT") = "") or (objRS("TXBLOK") = 0) Then
				Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			else
				Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			end if
		Else
			If (objRS("TXLOT") <> "") or (objRS("TXBLOK") > 0) Then
				Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			else
				Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
			end if
		End If
		intPrcl = objRS("TXPRCL")

	Case 4
		If  objRS("TXFLAG")= "V" then
			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=0&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
			Response.Write("<td class='rText'nowrap width='100'>"& strFlag21 &"</td>")
			Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
			response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
			If  objRS("TXSECT") > 0 Then
				If (objRS("TXLOT") = "") or (objRS("TXBLOK") = 0) Then
					Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				else
					Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				end if
			Else
				If (objRS("TXLOT") <> "") or (objRS("TXBLOK") > 0) Then
					Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				else
					Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				end if
			End If
			intPrcl = objRS("TXPRCL")
		else
			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
			Response.Write("<td class='rText'nowrap width='100'>"& strFlag21 &"</td>")
			Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
			response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
			If  objRS("TXSECT") > 0 Then
				If (objRS("TXLOT") = "") or (objRS("TXBLOK") = 0) Then
					Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				else
					Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				end if
			Else
				If (objRS("TXLOT") <> "") or (objRS("TXBLOK") > 0) Then
					Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				else
					Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
				end if
			End If
		end if
		intPrcl = objRS("TXPRCL")

	Case 5
		if intPrcl = objRS("TXPRCL") then
			if objRS("TXFLAG")= "T" Then
				Response.Write("<tr><td class='rText'><a class='pLink'nowrap href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
				Response.Write("<td class='rText'nowrap width='100'>"& strFlag21 &"</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
				If  objRS("TXSECT") > 0 Then
					If (objRS("TXLOT") = "") or (objRS("TXBLOK") = 0) Then
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				Else
					If (objRS("TXLOT") <> "") or (objRS("TXBLOK") > 0) Then
						Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				End If
			end if
			if objRS("TXFLAG")= "V" Then
				Response.Write("<tr><td class='rText'><a class='pLink'nowrap href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=1&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
				Response.Write("<td class='rText'nowrap width='100'>"& strFlag21 &"</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
				If  objRS("TXSECT") > 0 Then
					If (objRS("TXLOT") = "") or (objRS("TXBLOK") = 0) Then
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				Else
					If (objRS("TXLOT") <> "") or (objRS("TXBLOK") > 0) Then
						Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				End If
			end if
			intPrcl = objRS("TXPRCL")
		else
			if objRS("TXFLAG")= "T" Then
				Response.Write("<tr><td class='rText'><a class='pLink'nowrap href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
				Response.Write("<td class='rText'nowrap width='100'>"& strFlag21 &"</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
				If  objRS("TXSECT") > 0 Then
					If (objRS("TXLOT") = "") or (objRS("TXBLOK") = 0) Then
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				Else
					If (objRS("TXLOT") <> "") or (objRS("TXBLOK") > 0) Then
						Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				End If
			end if
			if objRS("TXFLAG")= "V" Then
				Response.Write("<tr><td class='rText'><a class='pLink'nowrap href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
				Response.Write("<td class='rText'nowrap width='100'>"& strFlag21 &"</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
				If  objRS("TXSECT") > 0 Then
					If (objRS("TXLOT") = "") or (objRS("TXBLOK") = 0) Then
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>S/T/R " & objRS("TXSECT") & "-" & objRS("TXTOWN") & "-" & objRS("TXRANG") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				Else
					If (objRS("TXLOT") <> "") or (objRS("TXBLOK") > 0) Then
						Response.Write("<td class='rText'nowrap>L/B/P " & objRS("TXLOT") & "-" & objRS("TXBLOK") & " " & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					else
						Response.Write("<td class='rText'nowrap>" & objRS("TXPLATD") &  " " & objRS("TXDSC1") & "</td>")
					end if
				End If
			end if
			intPrcl = objRS("TXPRCL")
		End If
	End Select

	objRS.MoveNext

Wend

objRS.Close
Set objRS = Nothing

%>
</table>
</body>
</html>