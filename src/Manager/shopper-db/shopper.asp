<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   SHOPPER.ASP                                                             %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% REM   header: %>
<% pageTitle = "Shopper Manager" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<% REM   body: %>
<UL>
    <LI> <A HREF="shopper_list.asp">  All Shoppers </A>
    <LI> <A HREF="shopper_day.asp">   New Shoppers by Month </A>
    <LI> <A HREF="shopper_month.asp"> New Shoppers by Year </A>
</UL>

<% REM   footer: %>
<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
