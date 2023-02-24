<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Notice of Disclaimer</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
	Dim cid, cName
	cid = request.QueryString("CID")

	If cid = "" Then
		response.Redirect("tax.asp")
	End If

	Select Case cid
		Case 21
		'Douglas County
		cName = "Douglas"
		Case 67
		'Rock County
		cName = "Rock"
		Case 53
		'Nobles County
		cName = "Nobles"
		Case 54
		'Norman County
		cName = "Norman"
		Case 65
		'Renville County
		cName = "Renville"
		Case 75
		'Stevens County
		cName = "Stevens"
		Case 76
		'Swift County
		cName = "Swift"
		Case 87
		'Yellow Medicine County
		cName = "Yellow Medicine"
		Case 61
		'Pope County
		cName = "Pope"
		Case 34
		'Kandiyohi County
		cName = "Kandiyohi"
		Case 47
		'Meeker County
		cName = "Meeker"
		Case Else
		'Kandiyohi County
		cName = "Kandiyohi"
	End Select
	Session("CountyID")= cid
	Session("County Name")= cName
	response.Write("<link rel='stylesheet' href='" & cid & ".css' type='text/css'>")
%>
</head>

<body>
<%

'Dim objCommand, objRScount, intnumberQ

'Set objCommand = Server.CreateObject("ADODB.Command")


'Fill in the command properties'
'Response.Write( "the strConnectCNT is " & strConnectCNT   )
'intnumberQ = 1
'response.Write(" the cid is " & cid  )
'response.Write(" the session county name is : " & cName )
'response.Write("the value of intnumberQ : " & intnumberQ )
'intnumberQ = "1"
'strQueryString = "WHERE [Counter" & cid & "].ID = '" & intnumberQ & "' "
'Response.Write("the value of the query string :" & strQueryString )
'objCommand.ActiveConnection = strConnectCNT
'strQueryString = "SELECT * FROM [Counter" & cid & "] WHERE [Counter" & cid & "].count ='1'"
'objCommand.CommandText = "SELECT * FROM [Counter" & cid & "] WHERE [Counter" & cid & "].count ='1'"
'Response.Write(" begin the value of the SELECT part of the Query String : " & strQueryString )
'objCommand.CommandType = 1
'Set objRScount = objCommand.Execute
'
'Set objCommand = Nothing






'Dim objCommand, objRScount, numberQ

'Set objCommand = Server.CreateObject("ADODB.Command")


'Fill in the command properties'
'Response.Write( "the strConnectCNT is " & strConnectCNT   )
'numberQ = 1
'response.Write(" the cid is " & cid  )
'response.Write("the value of NUMBER : " & numberQ )
'numberQ = 1
'strQueryString = "WHERE [Table1visitors].ID = '" & numberQ & "' "
'Response.Write("the value of the query string :" & strQueryString )
'objCommand.ActiveConnection = strConnectCNT
'strQueryString = "SELECT * FROM [Table1visitors] WHERE [Table1visitors].ID = '" & numberQ & "' "
'objCommand.CommandText = "SELECT * FROM [Table1visitors] WHERE [Table1visitors].ID = '" & numberQ & "' "
'Response.Write(" the value of the SELECT part of the Query String : " & strQueryString )
'objCommand.CommandType = 1
'Set objRScount = objCommand.Execute

'Set objCommand = Nothing

%>

	<!--#include file="PageCounters.asp"-->



<p>


	<table border="0" cellspacing="0" cellpadding="0" align="center">
		<tr valign="top">
			<td width="450" align="justify" class="STitle">Notice of Disclaimer:</td>
		</tr>
		<tr>
			<td width="450" bgcolor="#000000"></td>
		</tr>
		<tr>
			<td width="450" align="justify" class="rText">
			<% Response.Write(cName & " County makes the Web information available on an &quot;as is&quot; basis.  All warranties and representations of any kind with regard to said information is disclaimed, including the implied warranties of merchantability and fitness for a particular use.  " & cName & " County does not warrant the information against deficiencies of any kind.  Under no circumstances will " & cName & " County, or any of its officers or employees be liable for any consequential, incidental, special or exemplary damages even if appraised of the likelihood of such damages occurring.</td>") %>
		</tr>
	</table>
<form action="searchinput_Rock.asp" method="post">
<center>
<input type="radio" name="decision" value="iAccept">I ACCEPT the terms and conditions provided within this disclaimer.<br>
<input type="radio" name="decision" value="iDecline" checked>I DO NOT AGREE to the terms and conditions provided here.<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <br>
<input name="decisionButton" type="submit" value="Proceed">&nbsp;&nbsp;&nbsp;
</center>

</form>
<center>
<%
	'If cName = "Douglas" Then
	'	Response.Write("<input name='returnHomeButton' type='button' value='Return to Douglas Home Page' onClick=window.location='http://www.co.douglas.mn.us'>")
	'End If

	'Response.Write("You are the " & numberQ & " visitor to the " & cName & " County tax information site. ")
%>


<%
'Response.Write("Your session started at : " & Session("Start") )

'Response.Write("there have been " & Session("VisitorID") & " total visits to this site")

'Dim number
'Response.Write("value of COUNT : ")
'Response.Write(  objRScount("COUNT") )
'Response.Write(  objRScount("KEYID") )
'Response.Write(objRScount("TOTCOUNT"))
'number = objRScount("TOTCOUNT")
'number = objRScount("count")
'Response.Write("the value of number the first time is : " & number )
'objRScount("TOTCOUNT")  = (objRScount("TOTCOUNT")) + 1
'Response.Write("the value of number the second time is : " & number )
'objRScount.update
'number = objRScount("COUNT")
'objRScount.Close
'Set objRScount = Nothing






%>

</center>
</body>
</html>
