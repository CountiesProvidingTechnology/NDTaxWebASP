<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Notice of Disclaimer</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%

	Dim varcid, cName
	varcid = request.QueryString("cid")
'response.write(" the varcid is :" & varcid)
	If varcid = "" Then
		response.Redirect("tax.asp")
	End If

	Select Case varcid
		Case 21
		'Douglas County
		cName = "Douglas"
		Case 26
		'Grant County
		cName = "Grant"
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
		Case 74
		'Steele County
		cName = "Steele"
		Case 48
		'Mille Lacs County
		cName = "Mille Lacs"
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
		Case 42
		'Lyon County
		cName = "Lyon"
		Case 78
		'Traverse County
		cName = "Traverse"
		Case 84
		'Wilkin County
		cName = "Wilkin"
		Case Else
		'Kandiyohi County
		cName = "Kandiyohi"
	End Select
	Session("CountyID")= cid
	Session("County Name")= cName
	response.Write("<link rel='stylesheet' href='" & varcid & ".css' type='text/css'>")
%>
</head>

<body>



<!-- #include file="PageCounters.asp" -->

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
<form action="Searchinput_value.asp" method="post">
<center>
<input type="radio" name="decision" value="iAccept">I ACCEPT the terms and conditions provided within this disclaimer.<br>
<input type="radio" name="decision" value="iDecline" checked>I DO NOT AGREE to the terms and conditions provided here.<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <br>
<input type="hidden" name="cid" value=<%=varcid%> >
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
