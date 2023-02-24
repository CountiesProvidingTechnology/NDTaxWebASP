

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

          <td width="240" class="STitle">Taxpayer #<%= objRSPY("TXTAXP") %></td>
								</tr>
								<td width="240" class="rText">
									<%= objRSPY("TXTNAM") %><br>
									<%= objRSPY("TXTAD1") %><br>
									<%= objRSPY("TXTAD2") %><br>
									<%= objRSPY("TXTAD3") %><br>
									<%= objRSPY("TXTAD4") %><br>
								<tr valign="top">
									<td width="240"></td>
								</tr>
								<tr valign="top">
			<% If objRSPY("TXALTR") > 0 Then %>
          <td width="240" class="STitle">Alternate Taxpayer #<%= objRSPY("TXALTR") %></td>
								</tr>
								<%
									If objRSPY("TXALTR") > 0 Then
										Response.Write("<td width='240' class='rText'>")
										Response.Write(objRSPY("TXANAM") & "<br>")
										Response.Write(objRSPY("TXAAD1") & "<br>")
										Response.Write(objRSPY("TXAAD2") & "<br>")
										Response.Write(objRSPY("TXAAD3") & "<br>")
										Response.Write(objRSPY("TXAAD4") & "</td>")
									End If
								%>
				<% Else %>
					<td width="240" class="STitle"></td>
				<% End If %>
				<% If objRSPY("TXOWNR#") > 0 Then %>
								<tr valign="top">
									<td width="240" class="STitle">Owner #<%= objRSPY("TXOWNR#") %></td>
								</tr>
								<%
									If objRSPY("TXOWNR#") > 0 Then
										Response.Write("<td width='240' class='rText'>")
										Response.Write(objRSPY("TXONAM") & "<br>")
										Response.Write(objRSPY("TXOAD1") & "<br>")
										Response.Write(objRSPY("TXOAD2") & "<br>")
										Response.Write(objRSPY("TXOAD3") & "<br>")
										Response.Write(objRSPY("TXOAD4") & "</td>")
									End If
								%>
				<% Else %>
								<tr valign="top">
									<td width="240" class="STitle"></td>
								</tr>
				<% End If %>
				<% If objRSPY("TXFALC") > 0 Then %>
								<tr valign="top">
									<td width="240" class="STitle">Falco # <%= objRSPY("TXFALC") %></td>
								</tr>
								<%
									If objRSPY("TXFALC") > 0 Then
										Response.Write("<td width='240' class='rText'>")
										Response.Write(objRSPY("TXFALD") & "<br>")
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
										<td width="375" class="STitle">MP #<%= objRSPY("TXMP#") %>&nbsp;&nbsp;Re/Mh: <font class="oText"><%= objRSPY("TXREMH") %></font>
											<table border="0" cellspacing="0" cellpadding="0">
												<tr valign="top">
													<td width="75" class="STitle" align="center">Twp/City</td>
													<td width="75" class="STitle" align="center">School</td>
													<% If cid = 21 Then  %>
													<td width="75" class="STitle" align="center">Sewer</td>
													<td width="75" class="STitle" align="center">Water Shed</td>
													<td width="75" class="STitle" align="center">Fire</td>
													<td width="75" class="STitle" align="center">City</td>
													<% End If %>
													<% If cid = 67 Then %>
													<td width="75" class="STitle" align="center">Water Shed</td>
													<td width="75" class="STitle" align="center">*****</td>
													<td width="75" class="STitle" align="center">*****</td>
													<td width="75" class="STitle" align="center">R/U</td>
													<% End If %>
													<% If cid = 75 Then %>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">Fire</td>
													<td width="75" class="STitle" align="center">HRA</td>
													<td width="75" class="STitle" align="center">Agri</td>
													<% End If %>
													<% If cid = 65 Then %>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">RUSR</td>
													<td width="75" class="STitle" align="center">Debt</td>
													<td width="75" class="STitle" align="center">HRA</td>
													<% End If %>
													<% If cid = 61 Then %>
													<td width="75" class="STitle" align="center">Hosp</td>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">Sani</td>
													<td width="75" class="STitle" align="center">*****</td>
													<% End If %>
													<% If cid = 54 Then %>
													<td width="75" class="STitle" align="center">Wtr</td>
													<td width="75" class="STitle" align="center">*****</td>
													<td width="75" class="STitle" align="center">*****</td>
													<td width="75" class="STitle" align="center">*****</td>
													<% End If %>
													<% If cid = 34 Then %>
													<td width="75" class="STitle" align="center">Debt</td>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">Sewer</td>
													<td width="75" class="STitle" align="center">Agri</td>
													<% End If %>
													<% If cid = 87 Then %>
													<td width="75" class="STitle" align="center">WS</td>
													<td width="75" class="STitle" align="center">HD</td>
													<td width="75" class="STitle" align="center">Debt</td>
													<td width="75" class="STitle" align="center">RSD</td>
													<% End If %>
													<% If cid = 76 Then %>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">Fire</td>
													<td width="75" class="STitle" align="center">Debt</td>
													<td width="75" class="STitle" align="center">Agri</td>
													<% End If %>
													<% If cid = 53 Then %>
													<td width="75" class="STitle" align="center">WRSD</td>
													<td width="75" class="STitle" align="center">HRA</td>
													<td width="75" class="STitle" align="center">****</td>
													<td width="75" class="STitle" align="center">R/U</td>
													<% End If %>
													<% If cid = 47 Then %>
													<td width="75" class="STitle" align="center">Water</td>
													<td width="75" class="STitle" align="center">SUBO</td>
													<td width="75" class="STitle" align="center">Debt</td>
													<td width="75" class="STitle" align="center">*****</td>
													<% End If %>

												</tr>
												<tr valign="top">
													<td width="75" align="center" class="rText"><%= objRSPY("TXCITY") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSCHL") %></td>
													<% If cid = 21 Then %>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD4") %></td>
													<% End If %>
													<% If cid = 67 Then %>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%'= objRSPY("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%'= objRSPY("") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD2") %></td>
													<% End If %>
													<% If cid = 75 Then %>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD4") %></td>
													<% End If %>
													<% If cid = 65 Then %>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD4") %></td>
													<% End If %>
													<% If cid = 61 Then %>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%'= objRSPY("TXSTD4") %></td>
													<% End If %>
													<% If cid = 54 Then %>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%'= objRSPY("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%'= objRSPY("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%'= objRSPY("TXSTD4") %></td>
													<% End If %>
													<% If cid = 34 Then %>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD4") %></td>
													<% End If %>
													<% If cid = 87 Then %>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD4") %></td>
													<% End If %>
													<% If cid = 76 Then %>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD4") %></td>
													<% End If %>
													<% If cid = 53 Then %>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%'= objRSPY("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD3") %></td>
													<% End If %>
													<% If cid = 47 Then %>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%'= objRSPY("TXSTD4") %></td>
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
													<td width="75" align="center" class="rText"><%= objRSPY("TXSECT") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXTOWN") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXRANG") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXLOT") %></td>
													<td width="75" align="center" class="rText"><%= objRSPY("TXBLOK") %></td>
												</tr>
											</table>
											<br><%= objRSPY("TXPLATD") %><br><%= objRSPY("TXDSC1") %><br><%= objRSPY("TXDSC2") %><br><%= objRSPY("TXDSC3") %><br><%= objRSPY("TXDSC4") %>
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
										If objRSPY("TXPADR1") <> "" Then

										Response.Write( objRSPY("TXPADR1"))
										End If
										%>&nbsp;&nbsp;
										<%
										If objRSPY("TXPADR1") <> "" Then
										If objRSPY("TXPZIP1") = "00000"  Then
										Response.Write(" ")
										Else
                                        If objRSPY("TXPZIP1") = "000000000" Then
                                        Response.Write(" ")
										Else
										'Response.Write("here at last")
										Response.Write(objRSPY("TXPZIP1"))
										End if
										End If
										End If
									    %>
										<%'=' objRSPY("TXPADR1") %>&nbsp;&nbsp;<%'=' objRSPY("TXPZIP1") %>
										</td>
									</tr>
									<tr>
										<td width="375" class="rText">
										<%
											If objRSPY("TXPADR2") <> "" Then
											Response.Write( objRSPY("TXPADR2"))
											End If
										%>&nbsp;&nbsp;
										<%
											If objRSPY("TXPADR1") <> "" Then
												If objRSPY("TXPZIP2") = "00000"  Then
												Response.Write(" ")
												Else
												If objRSPY("TXPZIP2") = "000000000" Then
												Response.Write(" ")
												Else
												'Response.Write("here at last")
												Response.Write(objRSPY("TXPZIP2"))
												End if
												End If
										End If
									    %>
										<%'=' objRSPY("TXPADR2") %>&nbsp;&nbsp;<%'=' objRSPY("TXPZIP2") %></td>
									</tr>
									<tr>
										<td width="375" class="rText">&nbsp;&nbsp;
										<%
										If objRSPY("TXPADR3") <> "" Then
											Response.Write( objRSPY("TXPADR3"))
										End If
									    %>&nbsp;&nbsp;
									    <%
										If objRSPY("TXPADR3") <> "" Then
											If objRSPY("TXPZIP3") = "00000"  Then
											Response.Write(" ")
											Else
											If objRSPY("TXPZIP3") = "000000000" Then
											Response.Write(" ")
											Else
											'Response.Write("here at last")
											Response.Write(objRSPY("TXPZIP3"))
											End if
											End If
										End If
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
										<td width="375" class="rText"><%= objRSPY("TXESCR") %>&nbsp;&nbsp;<%= objRSPY("TXENAM") %></td>
									</tr>
									<tr>
										<td width="375" class="STitle">Deeded Acres: <font class="oText"><%= FormatNumber(objRSV("DEED"), 2) %></font></td>
									</tr>
									<tr>
										<td width="375" class="sTitle"></td>
									</tr>

									</td>
								</table>
							</td>
						</tr>
					</table>