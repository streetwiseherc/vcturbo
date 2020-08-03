<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   ERROR.ASP                                                               %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% REM   header: %>
<% pageTitle = "Error" %>

<HTML>
<HEAD>
    <TITLE> <% = pageTitle %> </TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>
<BODY TOPMARGIN="<% = top_margin %>" LEFTMARGIN="<% = left_margin %>" BGCOLOR="<% = bgcolor %>" TEXT="<% = text %>" LINK="<% = color_link %>" ALINK="<% = color_alink %>" VLINK="<% = color_vlink %>">
<!--#INCLUDE FILE="mgmt_header.asp" -->

<% REM   body: %>
<P>
We are sorry, but we are unable to process your request:
<BR>
<UL>
    <%
    for each errorStr in errorList
        Response.Write "<LI>" & errorStr & "</LI>"
    next
    %>
</UL>

Please go back and correct the error and try again.

</BODY>
</HTML>
