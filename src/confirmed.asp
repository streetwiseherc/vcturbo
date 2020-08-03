<%@ LANGUAGE=vbscript enablesessionstate=false %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   CONFIRMED.ASP                                                          #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<HTML>

<HEAD>
    <TITLE>Purchase Confirmation</TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>">

<% hd_image = "heading_checkout.gif" %>
<% alt = "Checkout - Confirmation"%>
<!--#INCLUDE FILE = "include/header.asp" -->

<%' Text Table %>
<TABLE WIDTH=600>
    <TR>
        <TD>
        <H1>Purchase Confirmed</H1>

        Congratulations! Your order has been processed successfully. We'll be shipping out your order right away.
        <P>
        For future reference, please make a note of the tracking number for this order. <BR> Click on the number to view your receipt.
        <P>
        <img src="assets/images/trackno.gif" alt="Your Tracking Number is <% = Request("order_id") %>" valign=bottom>
        &nbsp;&nbsp;&nbsp;&nbsp;
        <B><A HREF="<% = baseURL("receipt_view.asp") & mscsPage.URLShopperArgs("order_id", Request("order_id")) %>"><% = Request("order_id") %></A></B>
        <P>
        Thank you for choosing <% = displayName %>!
        <P>
        <A HREF="<% = pageURL("listing.asp") %>"><IMG SRC="assets/images/btn_continue_shopping.gif" BORDER=0 ALT="Continue Shopping"></A>
        <br>
        </TD>
    </TR>
</TABLE>
<% conn.close %>
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>

