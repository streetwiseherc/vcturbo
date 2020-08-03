<%@ LANGUAGE=vbscript enablesessionstate=false %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   DEFAULT.ASP                                                            #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<HTML>

<HEAD>
    <TITLE><% = displayName %></TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" background="<% = background %>">

<% hd_image="heading_welcome.gif" %>
<% alt="Welcome" %>
<!--#INCLUDE FILE = "include/header.asp" -->

<TABLE WIDTH=600>
    <TR>
        <TD align=left>If you've already registered, click the cup to <A HREF = "<% = pageURL("welcome_lookup.asp") %>">enter the store.</A></TD>
    </TR>
    <TR>
        <TD align=center><A HREF = "<% = pageURL("welcome_lookup.asp") %>"><IMG SRC="assets/images/title_volcano_coffee.gif" BORDER=0 ALT="<% = displayName %>" ALIGN=center></A></TD>
    </TR>
    <TR>
        <TD align=right>If you are new to <% = displayName %>, you may <A HREF = "<% = pageURL("welcome_new.asp") %>">register now.</A></TD>
    </TR>
</TABLE>

<!--#INCLUDE FILE = "include/footer.asp"-->

</BODY>
</HTML>
