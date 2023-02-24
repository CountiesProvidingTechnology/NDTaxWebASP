

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

          <td width="240" class="STitle">Taxpayer #<%= objRSPYGEN("TXTAXP") %></td>
								</tr>
								<td width="240" class="rText">
									<%= objRSPYGEN("TPNREV") %><br>
									<%= objRSPYGEN("TPADR1") %><br>
									<%= objRSPYGEN("TPADR2") %><br>
									<%= objRSPYGEN("TPADR3") %><br>
									<%= objRSPYGEN("TPADR4") %><br>
								<tr valign="top">
									<td width="240"></td>
								</tr>
								<tr valign="top">
			<% If objRSPYGEN("TXALTR") > 0 Then %>
          <td width="240" class="STitle">Alternate Taxpayer #<%= objRSPYGEN("TXALTR") %></td>
								</tr>
								<%
									If objRSPYGEN("TXALTR") > 0 Then
										Response.Write("<td width='240' class='rText'>")
										Response.Write(objRSPYGEN("ALNREV") & "<br>")
										Response.Write(objRSPYGEN("ALADR1") & "<br>")
										Response.Write(objRSPYGEN("ALADR2") & "<br>")
										Response.Write(objRSPYGEN("ALADR3") & "<br>")
										Response.Write(objRSPYGEN("ALADR4") & "</td>")
									End If
								%>
				<% Else %>
					<td width="240" class="STitle"></td>
				<% End If %>
				<% If objRSPYGEN("TXOWNR#") > 0 Then %>
								<tr valign="top">
									<td width="240" class="STitle">Owner #<%= objRSPYGEN("TXOWNR#") %></td>
								</tr>
								<%
									If objRSPYGEN("TXOWNR#") > 0 Then
										Response.Write("<td width='240' class='rText'>")
										Response.Write(objRSPYGEN("SVNREV") & "<br>")
										Response.Write(objRSPYGEN("SVADR1") & "<br>")
										Response.Write(objRSPYGEN("SVADR2") & "<br>")
										Response.Write(objRSPYGEN("SVADR3") & "<br>")
										Response.Write(objRSPYGEN("SVADR4") & "</td>")
									End If
								%>
				<% Else %>
								<tr valign="top">
									<td width="240" class="STitle"></td>
								</tr>
				<% End If %>
				<% If objRSPYGEN("TXFALC") > 0 Then %>
								<tr valign="top">
									<td width="240" class="STitle">Falco # <%= objRSPYGEN("TXFALC") %></td>
								</tr>
								<%
									If objRSPYGEN("TXFALC") > 0 Then
										Response.Write("<td width='240' class='rText'>")
										Response.Write(objRSPYGEN("TXFALD") & "<br>")
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
										<td width="375" class="STitle">MP #<%= objRSPYGEN("TXMP#") %>&nbsp;&nbsp;Re/Mh: <font class="oText"><%= objRSPYGEN("TXREMH") %></font>
											<table border="0" cellspacing="0" cellpadding="0">
												<tr valign="top">
													<td width="75" class="STitle" align="center">Twp/City</td>
													<td width="75" class="STitle" align="center">School</td>
													<% If cid = 27 Then %>
														<td width="75" class="STitle" align="center">Fire</td>
														<td width="75" class="STitle" align="center">Soil</td>
														<td width="75" class="STitle" align="center">Misc</td>
														<td width="75" class="STitle" align="center">COMR</td>
													<% End If %>
													<% If cid = 47 Then %>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">*****</td>
													<td width="75" class="STitle" align="center">Fire</td>
													<td width="75" class="STitle" align="center">*****</td>
													<% End If %>
													<% If cid = 30 Then %>
													<td width="75" class="STitle" align="center">Fire</td>
													<td width="75" class="STitle" align="center">Park</td>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center"></td>
													<% End If %>
													<% If cid = 31 Then %>
													<td width="75" class="STitle" align="center">Fire</td>
													<td width="75" class="STitle" align="center">Misc</td>
													<td width="75" class="STitle" align="center">AMB</td>
													<td width="75" class="STitle" align="center">Soil</td>
													<% End If %>
													<% If cid = 23 Then %>
													<td width="75" class="STitle" align="center">Fire</td>
													<td width="75" class="STitle" align="center">Park</td>
													<td width="75" class="STitle" align="center"></td>
													<td width="75" class="STitle" align="center"></td>
													<% End If %>
													<% If cid = 37 Then %>
													<td width="75" class="STitle" align="center">Fire</td>
													<td width="75" class="STitle" align="center">Park</td>
													<td width="75" class="STitle" align="center"></td>
													<td width="75" class="STitle" align="center"></td>
													<% End If %>
													<% If cid = 13 Then %>
													<td width="75" class="STitle" align="center">Fire</td>
													<td width="75" class="STitle" align="center">Ambl</td>
													<td width="75" class="STitle" align="center"></td>
													<td width="75" class="STitle" align="center"></td>
													<% End If %>
													<% If cid = 41 Then %>
													<td width="75" class="STitle" align="center">Fire</td>
													<td width="75" class="STitle" align="center">Park</td>
													<td width="75" class="STitle" align="center"></td>
													<td width="75" class="STitle" align="center"></td>
													<% End If %>

												</tr>
												<tr valign="top">
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXCITY") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSCHL") %></td>
													<% If cid = 27 Then %>
														<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD1") %></td>
														<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD2") %></td>
														<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD3") %></td>
														<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD4") %></td>
													<% End If %>
													<% If cid = 47 Then %>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD1") %></td><%'I have had to change the order of special tax districts. The placement on the Iseries data is one and three. %>
													<td width="75" align="center" class="rText"></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD3") %></td>
													<td width="75" align="center" class="rText"></td>
													<% End If %>
													<% If cid = 30 Then %>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD3") %></td>
													<td width="75" align="center" class="rText"></td>
													<% End If %>
													<% If cid = 31 Then %>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD4") %></td>
													<% End If %>
													<% If cid = 37 Then %>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD2") %></td>
													<td width="75" align="center" class="rText"></td>
													<td width="75" align="center" class="rText"></td>
													<% End If %>
													<% If cid = 23 Then %>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD2") %></td>
													<td width="75" align="center" class="rText"></td>
													<td width="75" align="center" class="rText"></td>
													<% End If %>
													<% If cid = 13 Then %>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD2") %></td>
													<td width="75" align="center" class="rText"></td>
													<td width="75" align="center" class="rText"></td>
													<% End If %>
													<% If cid = 41 Then %>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSTD2") %></td>
													<td width="75" align="center" class="rText"></td>
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
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXSECT") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXTOWN") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXRANG") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXLOT") %></td>
													<td width="75" align="center" class="rText"><%= objRSPYGEN("TXBLOK") %></td>
												</tr>
											</table>
											<br><%= objRSPYGEN("TXPLATD") %><br><%= objRSPYGEN("TXDSC1") %><br><%= objRSPYGEN("TXDSC2") %><br><%= objRSPYGEN("TXDSC3") %><br><%= objRSPYGEN("TXDSC4") %>
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
										<%
											If objRSPYGEN("PROPADR") <> "" Then

											Response.Write( objRSPYGEN("PROPADR"))
											'End If
											End If
										%>&nbsp;&nbsp;
										<%
										'If objRSPYGEN("TXPADR1") <> "" Then
										'If objRSPYGEN("TXPZIP1") = "00000"  Then
										'Response.Write(" ")
										'Else
                                       ' If objRSPYGEN("TXPZIP1") = "000000000" Then
                                       ' Response.Write(" ")
										'Else
										''Response.Write("here at last")
										'Response.Write(objRSPYGEN("TXPZIP1"))
										'End if
										'End If
										'End If
									    %>
										<%'=' objRSPYGEN("TXPADR1") %>&nbsp;&nbsp;<%'=' objRSPYGEN("TXPZIP1") %>
										</td>
									</tr>
									<tr>
										<td width="375" class="rText">
										<%
										'	If objRSPYGEN("TXPADR2") <> "" Then
										'	Response.Write( objRSPYGEN("TXPADR2"))
										'	End If
										%>&nbsp;&nbsp;
										<%
										'	If objRSPYGEN("TXPADR2") <> "" Then
										'		If objRSPYGEN("TXPZIP2") = "00000"  Then
										'		Response.Write(" ")
										'		Else
										'		If objRSPYGEN("TXPZIP2") = "000000000" Then
										'		Response.Write(" ")
										'		Else
										'		'Response.Write("here at last")
										'		Response.Write(objRSPYGEN("TXPZIP2"))
										'		End if
										'		End If
										'End If
									    %>
										<%'=' objRSPYGEN("TXPADR2") %>&nbsp;&nbsp;<%'=' objRSPYGEN("TXPZIP2") %></td>
									</tr>
									<tr>
										<td width="375" class="rText">&nbsp;&nbsp;
										<%
										'If objRSPYGEN("TXPADR3") <> "" Then
										'	Response.Write( objRSPYGEN("TXPADR3"))
										'End If
									    %>&nbsp;&nbsp;
									    <%
										'If objRSPYGEN("TXPADR3") <> "" Then
										'	If objRSPYGEN("TXPZIP3") = "00000"  Then
										'	Response.Write(" ")
										'	Else
										'	If objRSPYGEN("TXPZIP3") = "000000000" Then
										'	Response.Write(" ")
										'	Else
										'	'Response.Write("here at last")
										'	Response.Write(objRSPYGEN("TXPZIP3"))
										'	End if
										'	End If
										'End If
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
										<td width="375" class="rText"><%= objRSPYGEN("TXESCR") %>&nbsp;&nbsp;<%= objRSPYGEN("TXENAM") %></td>
									</tr>
									<tr>
										<td width="375" class="STitle">Deeded Acres: <font class="oText"><%= calcZeroRSPYGEN(("TXDEED"), 2) %></font></td>
									</tr>
									<tr>
										<td width="375" class="sTitle"></td>
									</tr>

									</td>
								</table>
							</td>
						</tr>
					</table>