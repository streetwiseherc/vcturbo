<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   ORDER.ASP                                                               %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "Order Manager" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<UL>
    <LI> <A HREF="order_list.asp"> All Orders </A>
    <LI> <A HREF="order_day.asp"> Orders by Month </A>
    <LI> <A HREF="order_month.asp"> Orders by Year </A>
    <LI> <A HREF="order_product.asp"> Orders by Product </A>
    <LI> <A HREF="order_shopper.asp"> Orders by Shopper </A>
</UL>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
