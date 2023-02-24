<%
	Dim strConnectCNT, strDbaseCNT
	strDbaseCNT = "Count" & Session("CountyID") & ".mdb"
	strConnectCNT = "provider=Microsoft.jet.oledb.4.0;" & _
				 "Data Source = " & strDbaseCNT & ";" & _
				 "Persist Security Info = False"
%>
