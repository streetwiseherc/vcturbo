<!--#INCLUDE FILE="include/shop.asp" -->
<!--#INCLUDE FILE="include/order.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   RECEIPT_LIST.ASP                                                          %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<HTML>
<HEAD>
<TITLE>Purchase History</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>">

<% hd_image = "heading_history.gif" %>
<% alt = "Checkout - Receipt"%>
<!--#INCLUDE FILE = "include/header.asp" -->

<%
Dim sql
Dim rsReceipts

sql = "select order_id, created, total_total from vcturbo_receipt where shopper_id = '" & mscsShopperId & "'"
set rsReceipts = conn.Execute(sql)
%>

<TABLE CELLPADDING=3>
    <% if rsReceipts.EOF then %>
        	<TR><TD><B>No receipts found</B></TD></TR>
        <% else %>
	        <TR><TH ALIGN=LEFT>Order Id</TH><TH ALIGN=LEFT>Date</TH><TH ALIGN=LEFT>Total</TH></TR>
	        <%while not rsReceipts.EOF
	        %>
	        	<TR>
	        	    <TD><A HREF="receipt_view.asp?<% = mscsPage.URLShopperArgs("order_id", rsReceipts("order_id").Value) %>"><% = rsReceipts("order_id").Value %></A></TD>
	        	    <TD><% = rsReceipts("created").Value %></TD>
	        	    <TD><% = MSCSDataFunctions.Money(rsReceipts("total_total").Value) %></TD>
				</TR>
				<% rsReceipts.MoveNext
	        wend %>

    <% end if %>

</TABLE>

<% REM -- footer:  %>
<!--#INCLUDE FILE="include/footer.asp" -->


</BODY>
</HTML>


