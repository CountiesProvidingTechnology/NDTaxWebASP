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
intLot = Request.Form("lName")

'***  IF user enters a * in the search box the user is taken back to the disclaimer page. added 9-23-08 LEM
if strAddress = "*" or strSAddress = "*" then
    Response.Redirect("disclaimer_value.asp?cid=" & cid & "")
End If
If strName = "*" Then
    Response.Redirect("disclaimer_value.asp?cid=" & cid & "")
End If


'**  Function to format the entry of the parcel number search from the user
'**  this will allow the user to enter numbers only in the text field of the form.
'**  this is in ND web tax
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
'**********
'**********
'The check for a * at the end of a search string. lem 3-12-09
'This function will be used in the Name, Address, Parcel, and Plat search field.
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
'*********
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
'** call the function if there is something to work with.  * * * * * *    Call the fmtparcelnum  function
' before the call to the fmtparcelnum function I need to know if the info from the search box is formatted with a '-' or not.
' to find out if the info from the search box has a '-' or not run the   ChckHyphen function first.  LEM 9-19-08
'* * * * * * * * * * *'

if strParcelNo = "*"  then
    Response.Redirect("SearchInput.asp?cid=" & cid & "")
End if
' * * * * * * * * * *
	chkit2=ChckHyphen(strParcelNo, chkit)
	if chkit = 0 then
 	strParcelNo = fmtparcelnum(strParcelNo)
 	end if
	NewParcl =  ChckStar(strParcelNo, checkstr)
	strParcelNo = checkstr

	strParcelNo = replace(strParcelNo, "*", "%")
	strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL Like '" & strParcelNo & "' ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
	intSet = 1
	strTitle1 = "Parcel No."
	strTitle2 = "Year"
	strTitle3 = "E911 Address"
	strTitle4 = "Name"
	strTitle5 = "Description"
End IF

If strAddress <> "" or strSAddress <> "" Then
	if strAddress <> "" Then
	strSearch = strAddress
	End If
	If strSAddress <> "" Then
	strSearch = strSAddress
	End If
	newSearch = ChckStar(strSearch, checkstar)
	strAddress = checkstar
	strSAddress = checkstar

	strAddress = replace(strAddress, "*", "%")
	strSAdress = replace(strSAdress, "*", "%")
	strQueryString = "WHERE [Table 1 - Name/Addr/Desc/Tax/Recap Info].PROPADR Like '" & strAddress & "'"

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
	newSearch = ChckStar(strName, checkstaradd)
	strName = checkstaradd
	strName = replace(strName, "*", "%")
	strName = replace(strName, "'", "''")
	 strQueryString = "WHERE [NAMESEARCH].TPNAME Like '" & strName & "' ORDER BY [NAMESEARCH].TPNAME, [NAMESEARCH].TXPRCL;"
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

	strQueryString = "WHERE  [Table 1 - Name/Addr/Desc/Tax/Recap Info ].TXTOWN = " & intTwp & " AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXRANG = " & intRange & " AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXSECT = " & intSect & " ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
	else

	strQueryString = "WHERE  [Table 1 - Name/Addr/Desc/Tax/Recap Info ].TXTOWN = " & intTwp & " AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXRANG = " & intRange & " ORDER BY [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPRCL, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXYEAR, [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXFLAG;"
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
	strQueryString = "WHERE  [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPLATD Like '" & strPlat & "'"
	strPlat = replace(strPlat, "%", "*")
	intSet = 5
End If
If intBlock <> "" Then
	strQueryString = "WHERE  [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXBLOK = " & intBlock & ""
	intSet = 5
End If
If intBlock <> "" and intLot <> "" Then
	strQueryString = "WHERE  [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXBLOK = " & intBlock & " AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXLOT = " & intLot & ""
	intSet = 5
End If
If strPlat <> "" and intLot <> ""  Then
	strPlat = replace(strPlat, "*", "%")
	strQueryString = "WHERE  [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPLATD Like '" & strPlat & "' AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXLOT = " & intLot & ""
	strPlat = replace(strPlat, "%", "*")
	intSet = 5
End If
If strPlat <> "" and intBlock <> ""  Then
	strPlat = replace(strPlat, "*", "%")
	strQueryString = "WHERE  [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPLATD Like '" & strPlat & "' AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXBLOK = " & intBlock & ""
	strPlat = replace(strPlat, "%", "*")
	intSet = 5
End If

If strPlat <> "" and intBlock <> "" and intLot <> "" Then
	strPlat = replace(strPlat, "*", "%")
	strQueryString = "WHERE  [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXPLATD Like '" & strPlat & "' AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXLOT = " & intLot & " AND [Table 1 - Name/Addr/Desc/Tax/Recap Info].TXBLOK = " & intBlock & ""
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

'Fill in the command properties
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


Response.Write("<table width='1200'><th class='STitle'nowrap>" & strTitle1 & "</th><th class='STitle'>" & strTitle2 & "</th><th class='STitle'>" & strTitle3 & "</th><th class='STitle'>" & strTitle4 & "</th><th class='STitle'>" & strTitle5 & "</th>")

While Not objRS.EOF

Response.Flush

	If objRS("TXFLAG") = "T" then
		strFlag = " Tax"
		strFlag21 = objRS("TXYEAR") & " Tax"
	end if
	If objRS("TXFLAG") = "V" then
		strFlag = " Value"
		strFlag21 = objRS("TXYEAR") & " Value for Tax Payable " & (objRS("TXYEAR") + 1)
	end if

	Select Case intSet
	Case 1														'Search by Parcel
		If  objRS("TXFLAG")= "T" then
			Response.Write("<tr><td class='rText'><a nowrap class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
		Else
			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=1&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
		End if
		Response.Write("<td class='rText'nowrap width='190'>"& strFlag21 &"</td>")
		Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
		response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
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

	Case 2														'Search by Address
			If  objRS("TXFLAG") = "T" then
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=1&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("PROPADR") & "</a></td>")
			Else
				Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=1&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("PROPADR") & "</a></td>")
			End if
			Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
			Response.Write("<td class='rText'nowrap width='190'>"& strFlag21 &"</td>")
			Response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
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

	Case 3														'Search by Name
		If  objRS("TXFLAG") = "T" then
			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=1&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TPNAME") & "</a></td>")
		Else
			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=1&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "'>" & objRS("TPNAME") & "</a></td>")
		End if
		Response.Write("<td class='rText'nowrap>" & objRS("TXPRCL") & "</td>")
		Response.Write("<td class='rText'nowrap width='190'>"& strFlag21 &"</td>")
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

	Case 4
		If  objRS("TXFLAG")= "V" then
			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=1&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
			Response.Write("<td class='rText'nowrap width='190'>"& strFlag21 &"</td>")
			Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
			response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
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
			Response.Write("<tr><td class='rText'nowrap><a class='pLink' href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
			Response.Write("<td class='rText'nowrap width='190'>"& strFlag21 &"</td>")
			Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
			response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
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

	Case 5
		if intPrcl = objRS("TXPRCL") then
			if objRS("TXFLAG")= "T" Then
				Response.Write("<tr><td class='rText'><a class='pLink'nowrap href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
				Response.Write("<td class='rText'nowrap width='190'>"& strFlag21 &"</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
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
			if objRS("TXFLAG")= "V" Then
				Response.Write("<tr><td class='rText'><a class='pLink'nowrap href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&rid=1&yr=" & objRS("TXYEAR") &  "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
				Response.Write("<td class='rText'nowrap width='190'>"& strFlag21 &"</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
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
		else
			if objRS("TXFLAG")= "T" Then
				Response.Write("<tr><td class='rText'><a class='pLink'nowrap href='Parcel.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
				Response.Write("<td class='rText'nowrap width='190'>"& strFlag21 &"</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
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
			if objRS("TXFLAG")= "V" Then
				Response.Write("<tr><td class='rText'><a class='pLink'nowrap href='Parcel_value.asp?pid=" & objRS("TXPRCL") & "&tid=0&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "'>" & objRS("TXPRCL") & "</a></td>")
				Response.Write("<td class='rText'nowrap>" & objRS("PROPADR") & "</td>")
				Response.Write("<td class='rText'nowrap width='190'>"& strFlag21 &"</td>")
				Response.Write("<td class='rText'nowrap>" & objRS("TPNAME") & "</td>")
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
	End Select

	objRS.MoveNext

Wend

objRS.Close
Set objRS = Nothing

%>

</table>
</body>
</html>
