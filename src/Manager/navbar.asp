<%@ LANGUAGE = VBScript %>
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   NAVBAR.ASP                                                             %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>
<!--#INCLUDE FILE="include/manager.asp" -->
<HTML>
<BODY BGCOLOR="<% = bgcolor %>" TEXT="<% = text %>" LINK="<% = color_link %>" ALINK="<% = color_alink %>" VLINK="<% = color_vlink %>">
<% pageTitle = Server.HTMLEncode(Application("MSCSSite").DisplayName) + " Site Manager" %>

<TABLE BORDER="0" CELLPADDING="3" CELLSPACING="0" WIDTH="100%">
    <TR>
        <TD ALIGN="LEFT" BGCOLOR="#000000">
            <FONT FACE="Arial, sans-serif" COLOR="#FFFFFF" SIZE="+2"><B><% = pageTitle %></B></FONT>
        </TD>
    </TR>
</TABLE>

<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="2" WIDTH="100%">
    <TR>
        <TD ALIGN="CENTER" COLSPAN="9" VALIGN="BOTTOM">
            <IMG SRC="MSCS_Images/extras/dot_black.gif" WIDTH="100%" HEIGHT="2" VSPACE="0" ALIGN="BOTTOM">
        </TD>
    </TR>
    <TR>
        <TD ALIGN="CENTER">
            <A HREF="manager.asp" target="body"> <FONT FACE="Arial, sans-serif" SIZE="-1"><B>Manager</B></FONT></A> 
        </TD>
        <TD ALIGN="CENTER" WIDTH="8"><IMG SRC="MSCS_Images/extras/dot_black.gif" WIDTH="2" HEIGHT="20" VSPACE="0"></TD>
        <TD ALIGN="CENTER"> 
            <A HREF="catalog-vc/dept_list.asp" target="body"> <FONT FACE="Arial, sans-serif" SIZE="-1"><B>Departments</B></FONT></A> 
        </TD>
        <TD ALIGN="CENTER"> 
            <A HREF="catalog-vc/product_list.asp" target="body"> <FONT FACE="Arial, sans-serif" SIZE="-1"><B>Products</B></FONT></A> 
        </TD>
        <TD ALIGN="CENTER" WIDTH="8"><IMG SRC="MSCS_Images/extras/dot_black.gif" WIDTH="2" HEIGHT="20" VSPACE="0"></TD>
        <TD ALIGN="CENTER"> 
            <A HREF="promo.asp" target="body"> <FONT FACE="Arial, sans-serif" SIZE="-1"><B>Promotions</B></FONT></A> 
        </TD>
        <TD ALIGN="CENTER" WIDTH="8" target="body"><IMG SRC="MSCS_Images/extras/dot_black.gif" WIDTH="2" HEIGHT="20" VSPACE="0"></TD>
        <TD ALIGN="CENTER"> 
            <A HREF="order-db/order.asp" target="body"> <FONT FACE="Arial, sans-serif" SIZE="-1"><B>Orders</B></FONT></A> 
        </TD>
        <TD ALIGN="CENTER"> 
            <A HREF="shopper-db/shopper.asp" target="body"> <FONT FACE="Arial, sans-serif" SIZE="-1"><B>Shoppers</B></FONT></A> 
        </TD>
    </TR> 
    <TR>
        <TD ALIGN="CENTER" COLSPAN="9" VALIGN="BOTTOM">
            <IMG SRC="MSCS_Images/extras/dot_black.gif" WIDTH="100%" HEIGHT="2" VSPACE="0" ALIGN="BOTTOM">
        </TD>
    </TR>
</TABLE>



</BODY>
</HTML>