<%@ LANGUAGE=vbscript %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   WELCOME_LOOKUP.ASP                                                     #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<HTML>

<HEAD>
    <TITLE>Lookup Shopper</TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" background="<% = background %>" onload="onload()">

<% hd_image="heading_welcome.gif" %>
<% alt="Welcome" %>
<!--#INCLUDE FILE = "include/header.asp" -->

<H2>Welcome back!</H2>
<P>
To start, enter your e-mail and password then click the "continue" button.
<P>
<SCRIPT LANGUAGE="Javascript">
    <!--
         function onload()
         {
             document.LookupUser.shopper_email.focus();
         }
     // -->
</SCRIPT>

<%' Logon table %>
<FORM METHOD=POST ACTION="<% = mscsPage.SURL("xt_shopper_lookup.asp") %>" NAME="LookupUser">
            <TABLE WIDTH=600 BORDER=0>
                <TR><TD>E-Mail:</TD>
                    <TD><INPUT TYPE="text" SIZE=32 NAME="shopper_email"></TD>
                </TR>
                <TR><TD> Password:</td>
                    <TD> <INPUT TYPE=password SIZE=32 NAME="shopper_password"></TD>                     
                </TR>
                <TR>
                    <TD WIDTH=430 COLSPAN=2 ALIGN=right>
                    <BR><BR><INPUT TYPE=image SRC="assets/images/btn_continue.gif" ALIGN=TOP BORDER=0>
                    </TD>
                </TR>
            </TABLE>
</FORM>
<% conn.close %>
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>

