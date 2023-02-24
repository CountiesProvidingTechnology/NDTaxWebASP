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
									<%= objRS("TPNREV") %><br>
									<%= objRS("TPADR1") %><br>
									<%= objRS("TPADR2") %><br>
									<%= objRS("TPADR3") %><br>
									<%= objRS("TPADR4") %><br>
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
										Response.Write(objRS("ALNREV") & "<br>")
										Response.Write(objRS("ALADR1") & "<br>")
										Response.Write(objRS("ALADR2") & "<br>")
										Response.Write(objRS("ALADR3") & "<br>")
										Response.Write(objRS("ALADR4") & "</td>")
									End If
								%>
				<% End If %>
				<% If objRS("TXOWNR#") > 0 Then %>
								<tr valign="top">
									<td width="240" class="STitle">Owner #<%= objRS("TXOWNR#") %></td>
								</tr>
								<%
									If objRS("TXOWNR#") > 0 Then
										Response.Write("<td width='240' class='rText'>")
										Response.Write(objRS("SVNREV") & "<br>")
										Response.Write(objRS("SVADR1") & "<br>")
										Response.Write(objRS("SVADR2") & "<br>")
										Response.Write(objRS("SVADR3") & "<br>")
										Response.Write(objRS("SVADR4") & "</td>")
									End If
								%>
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
													<td width="75" align="center" class="rText"><%= objRS("TXCITY") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSCHL") %></td>
													<% If cid = 27 Then %>
														<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
														<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
														<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
														<td width="75" align="center" class="rText"><%= objRS("TXSTD4") %></td>
													<% End If %>
													<% If cid = 47 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td><%'I have had to change the order of special tax districts. The placement on the Iseries data is one and three. %>
													<td width="75" align="center" class="rText"></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"></td>
													<% End If %>
													<% If cid = 30 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"></td>
													<% End If %>
													<% If cid = 31 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD3") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD4") %></td>
													<% End If %>
													<% If cid = 37 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"></td>
													<td width="75" align="center" class="rText"></td>
													<% End If %>
													<% If cid = 23 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"></td>
													<td width="75" align="center" class="rText"></td>
													<% End If %>
													<% If cid = 13 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
													<td width="75" align="center" class="rText"></td>
													<td width="75" align="center" class="rText"></td>
													<% End If %>
													<% If cid = 41 Then %>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD1") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXSTD2") %></td>
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
													<td width="75" align="center" class="rText"><%= objRS("TXSECT") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXTOWN") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXRANG") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXLOT") %></td>
													<td width="75" align="center" class="rText"><%= objRS("TXBLOK") %></td>
												</tr>
											</table>
											<br><%= objRS("TXPLATD") %><br><%= objRS("TXDSC1") %><br><%= objRS("TXDSC2") %><br><%= objRS("TXDSC3") %><br><%= objRS("TXDSC4") %>
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
										<% '
											If objRS("PROPADR") <> "" Then
											Response.Write( objRS("PROPADR"))
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