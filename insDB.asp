<%
	Dim strConnect, strDbase
	strDbase = "NDWebTab" & cid & ".mdb"
'	strDbase = "WebTabXX.mdb"

	strConnect = "provider=Microsoft.jet.oledb.4.0;" & _

				 "Data Source = \\192.168.1.17\Data\" & strDbase & ";" & _
				 "Persist Security Info = False"

%>