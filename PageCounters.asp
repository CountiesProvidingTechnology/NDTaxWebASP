<%
Dim x
Dim FSObject
Dim GetTotals
Dim WriteTotals
Dim PageEntry
Dim PageHits()
Dim PageName()
Dim ThisPage
%>

<!-- #include file="insDB.asp" -->

<%
Dim objCommand, objRSCNT, strQueryString, strPID, strTID, objRS3, objRS5, objRScount, intnumberQ

Set objCommand = Server.CreateObject("ADODB.Command")

objCommand.ActiveConnection = strConnect
objCommand.CommandText = "SELECT * FROM [Table 8 - Misc Data];"
objCommand.CommandType = 1

Set objRSCNT = objCommand.Execute

Set objCommand = Nothing

DisplayCount = objRSCNT("Counter")

If DisplayCount < 1 then
	DisplayCount = objRSCNT("CntrBkup")

	Dim  objBCKUP

	Set objCommand = Server.CreateObject("ADODB.Command")

	objCommand.ActiveConnection = strConnect
	objCommand.CommandText = "UPDATE [Table 8 - Misc Data] SET [Table 8 - Misc Data].[CntrBkup] = [CntrBkup]+1; "
	objCommand.CommandType = 1

	Set objBCKUP = objCommand.Execute

	Set objCommand = Nothing

	Set objCommand = Server.CreateObject("ADODB.Command")


	objCommand.ActiveConnection = strConnect
	objCommand.CommandText = "UPDATE [Table 8 - Misc Data] SET [Table 8 - Misc Data].[Counter] = [CntrBkup]+1; "
	objCommand.CommandType = 1

	Set objRSUDT = objCommand.Execute

	Set objCommand = Nothing
Else
	Dim  objRSUDT

	Set objCommand = Server.CreateObject("ADODB.Command")

	objCommand.ActiveConnection = strConnect
	objCommand.CommandText = "UPDATE [Table 8 - Misc Data] SET [Table 8 - Misc Data].[Counter] = [Counter]+1; "
	objCommand.CommandType = 1

	Set objRSUDT = objCommand.Execute

	Set objCommand = Nothing
	Response.Write("<font size=3><b>You are  visitor number...</b></font><br>" & DisplayCount & "</font>")
End if
%>