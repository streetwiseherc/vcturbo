<%@ LANGUAGE = vbscript %>
<% Response.ExpiresAbsolute=#Jan 01, 1980 00:00:00# %>

<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   CLOSED.ASP                                                              %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1997-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% REM   header: %>
<% pageTitle = "Store Closed" %>
<% Set mscsPage   = Server.CreateObject("Commerce.Page")%>
<% REM   from mgmt_header.asp %>
<HTML>
<HEAD>
    <TITLE> <% = pageTitle %> </TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">    
    <% REM   from mgmt_style.asp %>
    <STYLE>
    <!--
        BODY { text-align: left; color: black; font-family: Arial; font-size: 14pt }
        A:link { color: red; background: transparent; font-size: 10pt; text-decoration: none }
        TD.title { color: white; text-align: left; font-size: 18pt; font-weight: bold }
        H6 { text-align: left; color: black; font-family: Arial; font-size: 8pt font-weight: plain }
    -->
    </STYLE>
</HEAD>

<BODY TOPMARGIN=8 LEFTMARGIN=8 BGCOLOR=white>

<TABLE BORDER=0 CELLPADDING=3 CELLSPACING=0 WIDTH=100%>
    <TR>
        <TD CLASS=title ALIGN=LEFT BGCOLOR=black>
            <FONT COLOR=white><% = pageTitle %></FONT>
        </TD>
    </TR>
</TABLE>

<% REM   body: %>
<P>
We are sorry, but <STRONG><%= mscsPage.HTMLEncode(Application("MSCSSite").DisplayName) %></STRONG> is closed at the moment. Come again soon.

<% REM   from copyright.asp: %>
<P>
<H6>
    The Best Coffee on the Web
    <P>
    The names of companies, products, people, characters, and/or data mentioned herein are fictitious and are in no way intended to represent any real individual, company, product, or event, unless otherwise noted. 
    <P>
    &copy;1997-98 Microsoft Corporation, All rights reserved
</H6>
