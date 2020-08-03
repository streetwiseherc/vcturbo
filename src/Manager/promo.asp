<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PROMO.ASP                                                               %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% REM   header: %>
<% pageTitle = "Promo Manager" %>

<HTML>
<HEAD>
    <TITLE> <% = pageTitle %> </TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
    <% stdListRange = 15 %>
</HEAD>
<BODY TOPMARGIN="<% = top_margin %>" LEFTMARGIN="<% = left_margin %>" BGCOLOR="<% = bgcolor %>" TEXT="<% = text %>" LINK="<% = color_link %>" ALINK="<% = color_alink %>" VLINK="<% = color_vlink %>">
<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<% REM   body: %>
<UL>
    <LI><A HREF="promo-price/promo-price_list.asp"> Price Promotions </A>
    <LI><A HREF="promo-upsell/promo-upsell_list.asp"> Upsell Promotions </A>
    <LI><A HREF="promo-cross/promo-cross_list.asp"> Cross Promotions </A>
</UL>

</BODY>

</HTML>
