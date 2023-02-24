

<table border="0" cellspacing="0" cellpadding="0">
				<tr valign="top">
					<td width="275">

						<table border="0" cellspacing="0" cellpadding="0">
							<tr valign="top">
								<td width="240" class="SHeader">Taxpayer/Owner Information</td>
							</tr>
							<tr valign="top">
								<td width="240" bgcolor="#000000"></td>
							</tr>
							<tr valign="top">
								<tr valign="top">

          <td width="240" class="STitle">Taxpayer #<%= objRS("TXTAXP") %></td>
								</tr>
								<td width="240" class="rText">
									<%= objRS("TXTNAM") %><br>
									<%= objRS("TXTAD1") %><br>
									<%= objRS("TXTAD2") %><br>
									<%= objRS("TXTAD3") %><br>
									<%= objRS("TXTAD4") %><br>
								<tr valign="top">
									<td width="240"></td>
								</tr>
								<tr valign="top">
			<% If objRS("TXALTR") > 0 Then %>
          <td width="240" class="STitle">Alternate Taxpayer #<%= objRS("TXALTR") %></td>
								</tr>
								<%
									If objRS("TXALTR") > 0 Then
										Response.Write("<td width='240' class='rText'>")
										Response.Write(objRS("TXANAM") & "<br>")
										Response.Write(objRS("TXAAD1") & "<br>")
										Response.Write(objRS("TXAAD2") & "<br>")
										Response.Write(objRS("TXAAD3") & "<br>")
										Response.Write(objRS("TXAAD4") & "</td>")
									End If
								%>
				<% Else %>
					<td width="240" class="STitle"></td>
				<% End If %>
				<% If objRS("TXOWNR#") > 0 Then %>
								<tr valign="top">
									<td width="240" class="STitle">Owner #<%= objRS("TXOWNR#") %></td>
								</tr>
								<%
									If objRS("TXOWNR#") > 0 Then
										Response.Write("<td width='240' class='rText'>")
										Response.Write(objRS("TXONAM") & "<br>")
										Response.Write(objRS("TXOAD1") & "<br>")
										Response.Write(objRS("TXOAD2") & "<br>")
										Response.Write(objRS("TXOAD3") & "<br>")
										Response.Write(objRS("TXOAD4") & "</td>")
									End If
								%>
				<% Else %>
								<tr valign="top">
									<td width="240" class="STitle"></td>
								</tr>
				<% End If %>
				<% If objRS("TXFALC") > 0 Then %>
								<tr valign="top">
									<td width="240" class="STitle">Falco # <%= objRS("TXFALC") %></td>
								</tr>
								<%
									If objRS("TXFALC") > 0 Then
										Response.Write("<td width='240' class='rText'>")
										Response.Write(objRS("TXFALD") & "<br>")
									End If
								%>
				<% Else %>
								<tr valign="top">
									<td width="240" class="STitle"></td>
								</tr>
				<% End If %>
							</table>
							</td>
							<td width="375">

								<table border="0" cellspacing="0" cellpadding="0">
									<tr valign="top">
										<td width="375" class="SHeader">General</td>
									</tr>
									<tr valign="top">
										<td width="375" bgcolor="#000000"></td>
									</tr>
									<tr valign="top">
										<td width="375" class="STitle">MP #<%= objRS("TXMP#") %>&nbsp;&nbsp;Re/Mh: <font class="oText"><%= objRS("TXREMH") %></font>
											<table border="0" cellspacing="0" cellpadding="0">
												<tr valign="top">
													<td width="75" class="STitle" align="center">Twp/City</td>
													<td width="75" class="STitle" align="center">School</td>
													<% If varcid = 21 Then  %>
													<td width="75" class="STitle" align="center">Sewer</td>
													<td width="75" class="STitle" align="center">Water Shed</td>
													<td width="75" class="STitle" align="center">Fire</td>
													<td width="75" class="STitle" align="center">City</td>
													<% End If %>
													<% If varcid = 67 Then %>
													<td width="75" class="STitle" align="center">Water Shed</td>
													<td width="75" class="STitle" align="center">*****</td>
													<td width="75" class="STitle" align="center">*****</td>
													<td width="75" class="STitle" align="center">R/U</td>
													<% End If %>
													<% If varcid = 74 Then %>
													<td width="75" class="STitle" align="center">*****</td>
													<td width="75" class="STitle" align="center">*****</td>
													<td width="75" class="STitle" align="center">Debt</td>
													<td width="75" class="STitle" align="center">*****</td>
													<% End If %>
													<% If varcid = 75 Then %>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">Fire</td>
													<td width="75" class="STitle" align="center">HRA</td>
													<td width="75" class="STitle" align="center">Agri</td>
													<% End If %>
													<% If varcid = 65 Then %>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">RUSR</td>
													<td width="75" class="STitle" align="center">Debt</td>
													<td width="75" class="STitle" align="center">HRA</td>
													<% End If %>
													<% If varcid = 61 Then %>
													<td width="75" class="STitle" align="center">Hosp</td>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">Sani</td>
													<td width="75" class="STitle" align="center">*****</td>
													<% End If %>
													<% If varcid = 54 Then %>
													<td width="75" class="STitle" align="center">Wtr</td>
													<td width="75" class="STitle" align="center">*****</td>
													<td width="75" class="STitle" align="center">*****</td>
													<td width="75" class="STitle" align="center">*****</td>
													<% End If %>
													<% If varcid = 34 Then %>
													<td width="75" class="STitle" align="center">Debt</td>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">Sewer</td>
													<td width="75" class="STitle" align="center">Agri</td>
													<% End If %>
													<% If varcid = 87 Then %>
													<td width="75" class="STitle" align="center">WS</td>
													<td width="75" class="STitle" align="center">HD</td>
													<td width="75" class="STitle" align="center">Debt</td>
													<td width="75" class="STitle" align="center">RSD</td>
													<% End If %>
													<% If varcid = 76 Then %>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">Fire</td>
													<td width="75" class="STitle" align="center">Debt</td>
													<td width="75" class="STitle" align="center">Agri</td>
													<% End If %>
													<% If varcid = 77 Then %>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">****</td>
													<td width="75" class="STitle" align="center">Debt</td>
													<td width="75" class="STitle" align="center">Agri</td>
													<% End If %>
													<% If varcid = 53 Then %>
													<td width="75" class="STitle" align="center">WRSD</td>
													<td width="75" class="STitle" align="center">HRA</td>
													<td width="75" class="STitle" align="center">****</td>
													<td width="75" class="STitle" align="center">R/U</td>
													<% End If %>
													<% If varcid = 47 Then %>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">SUBO</td>
													<td width="75" class="STitle" align="center">Debt</td>
													<td width="75" class="STitle" align="center">*****</td>
													<% End If %>
													<% If varcid = 84 Then %>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">BRDR</td>
													<td width="75" class="STitle" align="center">*****</td>
													<td width="75" class="STitle" align="center">RSD</td>
													<% End If %>
													<% If varcid = 41 Then %>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">LBLI</td>
													<td width="75" class="STitle" align="center">Debt</td>
													<td width="75" class="STitle" align="center">*****</td>
													<% End If %>
												</tr>
												<tr valign="top">
													<td width="75" align="center" class="rText"><%= objRS("TXCITY") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSCHL") %></td>
													<% If varcid = 21 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD4") %></td>
													<% End If %>
													<% If varcid = 67 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%'= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%'= objRS("") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<% End If %>
													<% If varcid = 74 Then %>
													<td width="75" align="center" class="rText"></td>
													<td width="75" align="center" class="rText"></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"></td>
													<% End If %>
													<% If varcid = 75 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD4") %></td>
													<% End If %>
													<% If varcid = 65 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD4") %></td>
													<% End If %>
													<% If varcid = 61 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%'= objRS("TXSTD4") %></td>
													<% End If %>
													<% If varcid = 54 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%'= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%'= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%'= objRS("TXSTD4") %></td>
													<% End If %>
													<% If varcid = 34 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD4") %></td>
													<% End If %>
													<% If varcid = 87 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD4") %></td>
													<% End If %>
													<% If varcid = 76 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD4") %></td>
													<% End If %>
													<% If varcid = 77 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD4") %></td>
													<% End If %>
													<% If varcid = 53 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%'= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<% End If %>
													<% If varcid = 47 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%'= objRS("TXSTD4") %></td>
													<% End If %>
													<% If varcid = 84 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD4") %></td>
													<% End If %>
													<% If varcid = 41 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"></td>
													<% End If %>
												</tr>
											</table>
										</td>
									</tr>
									<tr valign="top">
										<td width="375"><br></td>
									</tr>
									<tr valign="top">
										<td width="375" class="SHeader">Description</td>
									</tr>
									<tr valign="top">
										<td width="375" bgcolor="#000000"></td>
									</tr>
									<tr valign="top">
										<td width="375" class="rText">
											<table border="0" cellspacing="0" cellpadding="0">
												<tr valign="top">
													<td width="75" align="center" class="STitle">Sect</td>
													<td width="75" align="center" class="STitle">Twp</td>
													<td width="75" align="center" class="STitle">Range</td>
													<td width="75" align="center" class="STitle">Lot</td>
													<td width="75" align="center" class="STitle">Block</td>
												</tr>
												<tr valign="top">
													<td width="75" align="center" class="rText"><%= objRS("TXSECT") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXTOWN") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXRANG") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXLOT") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXBLOK") %></td>
												</tr>
											</table>
											<br><%= objRS("TXPLATD") %><br><%= objRS("TXDSC1") %><br><%= objRS("TXDSC2") %><br><%= objRS("TXDSC3") %><br><%= objRS("TXDSC4") %><br><%= objRS("TXDSC5") %><br><%= objRS("TXDSC6") %><br><%= objRS("TXDSC7") %><br><%= objRS("TXDSC8") %>
						<%
							While Not ObjRS11.EOF %>
								 <br><%=  objRS11("TXDESC") %>

								<% objRS11.MoveNext
							Wend
						%>

										</td>
									</tr>
									<tr valign="top">
										<td width="375"><br></td>
									</tr>
									<tr valign="top">
										<td width="375" bgcolor="#000000"></td>
									</tr>
									<tr valign="top">
										<td width="375" class="STitle">Property Address</td>
									</tr>
									<tr>
										<td width="375" class="rText">
										<% 'If cid = 67 then
											If objRS("TXPADR1") <> "" Then

											Response.Write( objRS("TXPADR1"))
											'End If
											End If
											%>
											&nbsp;&nbsp;
											<%
											'If cid = 67 then
											If objRS("TXPADR1") <> "" Then
											If objRS("TXPZIP1") = "00000"  Then
											Response.Write(" ")
											Else
											If objRS("TXPZIP1") = "000000000" Then
											Response.Write(" ")
											Else
											'Response.Write("here at last")
											Response.Write(objRS("TXPZIP1"))
											End if
											End If
											End If
											'End If
											%>
										<%
										'If cid = 67 then
										'else

										'If objRS("TXPADR") <> "" Then
										'response.Write(objRS("TXPADR"))
										'Response.Write("  ")
										'If objRS("TXPZIP") = "00000"  Then
										'Response.Write(" ")
										'Else
										'If objRS("TXPZIP") = "0000000000" Then
										'Response.Write(" ")
										'Else
										''Response.Write("here at last")
										'Response.Write(objRS("TXPZIP"))
										'End if
										'End If
										'End If


										'response.Write(objRS("TXPADR"))
										'Response.Write("  ")
										'Response.Write(calcZip("TXPZIP"))
										'Response.Write(objRS("TXPZIP"))
										'end if
										%>
										<%'= calcZip("TXPZIP") %>
										<%'=' objRS("TXPADR") %>&nbsp;&nbsp;<%'=' objRS("TXPZIP") %>
									<% 'If objRS("TXPZIP") > 0 Then
											'calcZip("TXPZIP")
										'Else
									    'End If	'
									 %>
									</td>
									</tr>
									<tr>
										<td width="375" height="30"></td>
									</tr>
									<tr>
										<td width="375" class="STitle">Escrow</td>
									</tr>
									<tr>
										<td width="375" class="rText"><%= objRS("TXESCR") %>&nbsp;&nbsp;<%= objRS("TXENAM") %></td>
									</tr>
									<tr>
										<td width="375" class="STitle">Deeded Acres: <font class="oText"><%= calcZero(("TXDEED"), 2) %></font></td>
									</tr>
									<tr>
										<td width="375" class="sTitle"></td>
									</tr>

									</td>
								</table>
							</td>
						</tr>
					</table>