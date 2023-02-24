<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td width="600" colspan="1" class="sHeader">On-Line Payment</td>
	</tr>
	<tr valign="top">
		<td width="450" colspan="1" height="1" bgcolor="#000000"></td>
	</tr>
	<tr class="sTitle" valign="top">
		<td>Current amount due is...<%= FormatNumber(objRS5("UPTOTDUE1"), 2) %> </td>
	</tr>
	<tr class="sTitle" valign="top">
	<%
	
		payAmount = Session("payAmount")
		payAmount2= Request.Form("ccamount")
		If payAmount2>0 then
			payAmount=payAmount2
		end if
		if payAmount*1>objRS5("UPTOTDUE1") then
			payAmount=objRs5("UPTOTDUE1")	
		end if

		Dim lineData '= ""
		Set fso = Server.CreateObject("Scripting.FileSystemObject") 
		set fs = fso.OpenTextFile(Server.MapPath("HeartlandInfo.txt"), 1, true) 
		Do Until fs.AtEndOfStream 
   		 lineData =  lineData & fs.ReadLine
    		 'do some parsing on lineData to get image data
  		  'output parsed data to screen
   		 'Response.Write lineData
		Loop 

		fs.close: set fs = nothing 
		
	%>
		<td>Current amount of payment...<%= FormatNumber(payAmount*1,2) %> </td>
	</tr>
	<tr class="sTitle">
		<td>&nbsp;</td>
	</tr>
		<tr class="sTitle">
			<td>If you would like to change the amount enter it here and then click on Change Amount.  When you have the amount set, click on Submit Payment.  Note that you cannot pay more than the amount due.</td>
		</tr>
</table>

<%
	
	actionstr="Parcel.asp?pid=" & strPID & "&tid=7&cid=" & cid &"&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "" 
%>
<table border="0" cellpadding="0" cellspacing="0">
<form action=<%= actionstr %> method="post">
	<tr class="sTitle"><td width="200">&nbsp;</td><td width="200"></td><td width="300"></td></tr>
	<tr class="sTitle">
		<td width="200">
		<%	
		payAmount = Session("payAmount")
		
		payAmount2= Request.Form("ccamount")
		If payAmount2>0 then
			payAmount=payAmount2
		end if
		if payAmount*1>objRS5("UPTOTDUE1") then
			payAmount=objRs5("UPTOTDUE1")	
		end if
		payAmount=FormatNumber(payAmount,2)
		Session("pid") = strPID
		Session("cid") = cid
		Session("varintParcelNo")=intParcelNo
		Session("varstrAddress")=strAddress
		Session("varstrName")=strName
		Session("varintSect")=intSect
		Session("varintTwp")=intTwp
		Session("varintRange")=intRange

		%>
		
		<input type="text"  value=<%= payAmount %> name="ccamount"></td>
		<td width="200">
			<input type="submit" name="Button5" Value="Change Amount"></td>
		
		<td width="300">&nbsp;</td>
		
		</tr>
		</form>
</table>



<table border="0" cellpadding="0" cellspacing="0">
<tr class="sTitle"><td width="900">&nbsp;</td></tr>
	<tr class="sTitle" valign="top">
		<td width="900">
		<%

		//Dim strAmount
		payAmount = Session("payAmount")
		payAmount2= Request.Form("ccamount")
//response.Write (" Session Var " & payAmount  )
		If payAmount2>0 then
			payAmount=payAmount2
		end if
		
		if (payAmount*1>objRS5("UPTOTDUE1")) then
			payAmount=objRS5("UPTOTDUE1")	
		end if
		//payAmount=FormatNumber(payAmount,2)
 		Session("payAmount") = payAmount
		
		//cid=Session("cid")
		//intParcelNo=Session("varintParcelNo")
		//strAddress=Session("varstrAddress")
		//strName=Session("varstrName")
	//	intSect=Session("varintSect")
		//intTwp=Session("varintTwp")
	//	intRange=Session("varintRange")
payamount = replace(payamount,",","")

		//response.Write (" Session Var " & payAmount  )
		Dim objHTTP2, strEnvelope2
        			Set objHTTP2 = Server.CreateObject("MSXML2.ServerXMLHTTP")
					strEnvelope2 = "<?xml version='1.0' encoding='utf-8'?>"
        			strEnvelope2 = strEnvelope2 & "<s:Envelope xmlns:s='http://schemas.xmlsoap.org/soap/envelope/' xmlns:a='http://schemas.datacontract.org/2004/07/BDMS.NewModel' xmlns:b='https://test.heartlandpaymentservices.net/BillingDataManagement/v3/BillingDataManagementService'>"
					strEnvelope2 = strEnvelope2 & "<s:Body>"
					strEnvelope2 = strEnvelope2 & "<b:LoadSecurePayMerchantBillData >"
					strEnvelope2 = strEnvelope2 & "<b:request>"
					strEnvelope2 = strEnvelope2 & "<a:Credential>"
          			strEnvelope2 = strEnvelope2 & lineData
        			strEnvelope2 = strEnvelope2 & "</a:Credential>"
					strEnvelope2 = strEnvelope2 & "<a:BillData>"
          			strEnvelope2 = strEnvelope2 & "<a:SecurePayBill>"
            		strEnvelope2 = strEnvelope2 & "<a:Amount>" & payAmount & "</a:Amount>"
            		strEnvelope2 = strEnvelope2 & "<a:BillTypeName>Tax Payment</a:BillTypeName>"
            		strEnvelope2 = strEnvelope2 & "<a:Identifier1>" & strPID & "</a:Identifier1>"
          			strEnvelope2 = strEnvelope2 & "</a:SecurePayBill>"
		 			strEnvelope2 = strEnvelope2 & "</a:BillData>"
        			strEnvelope2 = strEnvelope2 & "<a:MaxFuturePaymentDays>0</a:MaxFuturePaymentDays>"
        			strEnvelope2 = strEnvelope2 & "<a:ReturnTokenWithResponse>false</a:ReturnTokenWithResponse>"
      				strEnvelope2 = strEnvelope2 & "</b:request>"
    				strEnvelope2 = strEnvelope2 & "</b:LoadSecurePayMerchantBillData>"
  					strEnvelope2 = strEnvelope2 & "</s:Body>"
					strEnvelope2 = strEnvelope2 & "</s:Envelope>" 
					//response.Write ("Envelope " & strEnvelope2  )
					Dim url2
        			url2 = "https://heartlandpaymentservices.net/BillingDataManagement/v3/BillingDataManagementService.svc"
					With objHTTP2
//response.Write ("Envelope22 " & url2  )
            			.Open "POST", url2, False
            			.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
            			.setRequestHeader "SOAPAction", "https://test.heartlandpaymentservices.net/BillingDataManagement/v3/BillingDataManagementService/IBillingDataManagementService/LoadSecurePayMerchantBillData"
            			.send strEnvelope2
//response.Write( objHTTP2.responseXML.Text)
        			End With
//response.Write ("Envelope2 "   )
					 Dim strResponse2
        			strResponse2 = objHTTP2.responseXML.Text	
					//response.Write("response=" & strResponse2)	
					Dim WordArray2
					WordArray2 = split(strResponse2, "true")
	
		If payAmount>0 then
			Response.Write("<a href='https://heartlandpaymentservices.net/SecurePay/MountrailCountyTaxCollector/" & WordArray2(1) & "' class='tLink2'  align='right'>&nbsp;&nbsp;Submit Payment of " & payAmount & "</a>&nbsp;&nbsp;&nbsp; ")
		end if
		Response.Write("<a href='Parcel.asp?pid=" & strPID & "&tid=1&cid=" & cid & "&varintParcelNo=" & intParcelNo & "&varstrAddress=" & strAddress & "&varstrName=" & strName & "&varintSect=" & intSect & "&varintTwp=" & intTwp & "&varintRange=" & intRange & "payAmount=" & payAmount & "' class='tLink2'>&nbsp;&nbsp;&nbsp;Cancel</a>   ")
		//Response.Write("<a href='https://testing.heartlandpaymentservices.net/SecurePay/ComputerProfessionalsUnlimitedtest/" & WordArray2(1) & "' class='tLink2'  align='right'>&nbsp;&nbsp;Electronic Payment</a> ")
		%>
		</td>
		</tr>
</table>